#!/usr/bin/env php
<?php

require __DIR__ . "/vendor/autoload.php";

use Jpastoor\JiraWorklogExtractor\Command\ClearCacheCommand;
use Jpastoor\JiraWorklogExtractor\Command\LoadProjectTotalsInTimeperiodCommand;
use Jpastoor\JiraWorklogExtractor\Command\WorkedHoursPerDayCommand;
use Jpastoor\JiraWorklogExtractor\Command\WorkedHoursPerDayPerAuthorCommand;
use Jpastoor\JiraWorklogExtractor\Command\WorklogEntriesCommand;
use Symfony\Component\Console\Application;

$application = new Application();
$application->add(new LoadProjectTotalsInTimeperiodCommand());
$application->add(new WorkedHoursPerDayCommand());
$application->add(new WorkedHoursPerDayPerAuthorCommand());
$application->add(new WorklogEntriesCommand());
$application->add(new ClearCacheCommand());
$application->run();
