<?php


namespace Jpastoor\JiraWorklogExtractor\Command;

use chobie\Jira\Api;
use Jpastoor\JiraWorklogExtractor\CachedHttpClient;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Helper\ProgressBar;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use XLSXWriter;

/**
 * Class WorkedHoursPerDayCommand
 *
 * Days on the rows
 * - columns: authors
 * - tabs: labels
 *
 * @package Jpastoor\JiraWorklogExtractor
 * @author Joost Pastoor <joost.pastoor@munisense.com>
 * @copyright Copyright (c) 2016, Munisense BV
 */
class WorklogEntriesCommand extends Command
{
    const MAX_ISSUES_PER_QUERY = 100;

    protected function configure()
    {
        $this
            ->setName('worklog-entries')
            ->setDescription('Exports worklog entries for all projects to excel.')
            ->addArgument(
                'start_time',
                InputArgument::REQUIRED,
                'From when do you want to load the worklog totals (YYYY-mm-dd)'
            )
            ->addArgument(
                'end_time',
                InputArgument::OPTIONAL,
                'End time to load the worklog totals (YYYY-mm-dd)',
                date("Y-m-d")
            )->addOption(
                'clear_cache', "c",
                InputOption::VALUE_NONE,
                'Whether or not to clear the cache before starting'
            )->addOption(
                'output-file', "o",
                InputOption::VALUE_REQUIRED,
                'Path to Excel file',
                __DIR__ . "/../../output/output_" . date("YmdHis") . ".xlsx"
            )->addOption(
                'authors-whitelist', null,
                InputOption::VALUE_OPTIONAL,
                'Whitelist of authors (comma separated)'
            )->addOption(
                'labels-whitelist', null,
                InputOption::VALUE_OPTIONAL,
                'Whitelist of labels (comma separated)'
            )->addOption(
                'projects-whitelist', null,
                InputOption::VALUE_OPTIONAL,
                'Whitelist of projects (comma separated)'
            )->addOption(
                'labels-blacklist', null,
                InputOption::VALUE_OPTIONAL,
                'Blacklist of labels (comma separated)'
            )->addOption(
                'config-file', null,
                InputOption::VALUE_OPTIONAL,
                'Path to config file',
                __DIR__ . "/../../config.json"
            );
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $start_time = $input->getArgument('start_time');
        $end_time = $input->getArgument('end_time');
        $start_time_obj = \DateTime::createFromFormat("Y-m-d", $start_time);
        $end_time_obj = \DateTime::createFromFormat("Y-m-d", $end_time);
        $start_timestamp = mktime(0, 0, 0, $start_time_obj->format("m"), $start_time_obj->format("d"), $start_time_obj->format("Y"));
        $end_timestamp = mktime(23, 59, 59, $end_time_obj->format("m"), $end_time_obj->format("d"), $end_time_obj->format("Y"));

        if (!file_exists($input->getOption("config-file"))) {
            $output->writeln("<error>Could not find config file at " . $input->getOption("config-file") . "</error>");
            die();
        }

        $config = json_decode(file_get_contents($input->getOption("config-file")));

        $cached_client = new CachedHttpClient(new Api\Client\CurlClient());
        $jira = new Api(
            $config->jira->endpoint,
            new Api\Authentication\Basic($config->jira->user, $config->jira->password),
            $cached_client
        );

        if ($input->getOption("clear_cache")) {
            $cached_client->clear();
        }

        $progress = null;
        $offset = 0;
        
        $worklogs = array();

        do {

            $jql = "worklogDate <= " . $end_time . " and worklogDate >= " . $start_time . " and timespent > 0  and timeSpent < " . rand(1000000, 9000000) . " ";

            if ($input->getOption("labels-whitelist")) {
                $jql .= " and labels in (" . $input->getOption("labels-whitelist") . ")";
                $labels_whitelist = explode(",", $input->getOption("labels-whitelist"));
            }

            if ($input->getOption("labels-blacklist")) {
                $jql .= " and (labels not in (" . $input->getOption("labels-blacklist") . ") OR labels is EMPTY )";
            }

            if ($input->getOption("authors-whitelist")) {
                $jql .= " and worklogAuthor in (" . $input->getOption("authors-whitelist") . ")";
            }

            if ($input->getOption("projects-whitelist")) {
                $jql .= " and project in (" . $input->getOption("projects-whitelist") . ")";
            }

            $search_result = $jira->search($jql, $offset, self::MAX_ISSUES_PER_QUERY, "key,project,labels,summary");

            if ($progress == null) {
                /** @var ProgressBar $progress */
                $progress = new ProgressBar($output, $search_result->getTotal());
                $progress->start();
            }

            // For each issue in the result, fetch the full worklog
            $issues = $search_result->getIssues();
            foreach ($issues as $issue) {
                $fields = $issue->getFields();
                $project = $fields['Project']["key"];
                $labels = $fields['Labels'];
                $summary = $fields['Summary'];

                if (isset($labels_whitelist)) {
                    $labels = array_intersect($labels, $labels_whitelist);
                }

                if (count($labels) > 1) {
                    $output->write("<error>" . $issue . " has multiple labels: " . implode(", ", $labels) . "</error>");
                }

                $worklog_result = $jira->getWorklogs($issue->getKey(), []);

                $worklog_array = $worklog_result->getResult();
                if (isset($worklog_array["worklogs"]) && !empty($worklog_array["worklogs"])) {
                    foreach ($worklog_array["worklogs"] as $entry) {
                        $author = $entry["author"]["key"];
                        // Filter on author
                        if ($input->getOption("authors-whitelist")) {
                            $authors_whitelist = explode(",", $input->getOption("authors-whitelist"));
                            if (!in_array($author, $authors_whitelist)) {
                                continue;
                            }
                        }

                        // Filter on time
                        $worklog_date = \DateTime::createFromFormat("Y-m-d", substr($entry['started'], 0, 10));
                        $worklog_timestamp = $worklog_date->getTimestamp();

                        if ($worklog_timestamp < $start_timestamp || $worklog_timestamp > $end_timestamp) {
                            continue;
                        }

                        $worklogs[$project][] = array(
                          'date' => $worklog_date->format("Y-m-d"),
                          'key' => $issue->getKey(),
                          'duration_m' => round($entry["timeSpentSeconds"] / 60),
                          'duration_h' => round(($entry["timeSpentSeconds"] / 3600), 2),
                          'author' => $author,
                          'labels' => implode(', ', $labels),
                          'summary' => $summary,
                          'comment' => $entry['comment'],
                        );
                    }
                }
                $progress->advance();
            }

            $offset += count($issues);
        } while ($search_result && $offset < $search_result->getTotal());

        $progress->finish();
        $progress->clear();

        if (empty($worklogs)) {
            throw new \Exception("No matching issues found");
        }

        $writer = new XLSXWriter();
        $writer->setAuthor("acolono GmbH");

        ksort($worklogs);

        $header = array(
          'Date' => 'date',
          'Issue' => 'string',
          'Duration (m)' => 'integer',
          'Duration (h)' => '#,##0.00',
          'Author' => 'string',
          'Labels' => 'string',
          'Summary' => 'string',
          'Comment' => 'string',
        );

        $sum_cells = array(2, 3);
        $projects = array_keys($worklogs);
        foreach ($projects as $project) {

            // Set Header captions.
            $writer->writeSheetHeader($project, $header);

            // Add rows with worklogs.
            $i = 0;
            foreach (array_reverse($worklogs[$project]) as $row) {
                $writer->writeSheetRow($project, $row);
                $i++;
            }

            // Build totals.
            $sums = array();
            foreach (array_keys($header) as $cell => $caption) {
                if (in_array($cell, $sum_cells)) {
                    $sums[] = "=SUM(" . XLSXWriter::xlsCell(1, $cell) . ":" . XLSXWriter::xlsCell($i, $cell) . ")";
                } else {
                    $sums[] = "";
                }
            }
            $writer->writeSheetRow($project, $sums);
        }

        $writer->writeToFile($input->getOption("output-file"));
    }
}
