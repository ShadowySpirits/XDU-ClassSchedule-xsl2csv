<?php

require './vendor/autoload.php';
require './helper.php';

use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\IOFactory;

try {
    $startDate = explode('-', $_POST['date']);
    $spreadsheet = IOFactory::load($_FILES['file']['tmp_name']);
    $worksheet = $spreadsheet->getActiveSheet();
    $csv_title = array('Subject', 'StartDate', 'EndDate', 'StartTime', 'EndTime', 'Location', 'Description');
    $csv_filed = array();
    foreach ($worksheet->getRowIterator(2) as $row) {
        $data = array();
        $index = $row->getRowIndex();

        $name = $worksheet->getCell("B$index")->getValue();
        $day = $worksheet->getCell("F$index")->getValue() - 1;
        $startNum = $worksheet->getCell("G$index")->getValue();
        $endNum = $worksheet->getCell("H$index")->getValue();
        $teacherName = $worksheet->getCell("I$index")->getValue();
        $classNum = $worksheet->getCell("A$index")->getValue();
        $location = $worksheet->getCell("J$index")->getValue();
        $weeks = $worksheet->getCell("E$index")->getValue();

        $duration = explode(',', $weeks);
        $singleOrDouble = 0;
        if (false !== strpos($weeks, '单')) {
            $singleOrDouble = 1;
        } else if (false !== strpos($weeks, '双')) {
            $singleOrDouble = 2;
        }

        foreach ($duration as $value) {
            $matches = array();
            preg_match_all('/\d+/', $value, $matches);
            /** @noinspection TypeUnsafeComparisonInspection */
            @$maxWeek = $matches[0][1] != 0 ? $matches[0][1] : $matches[0][0];
            for ($i = $matches[0][0] - 1; $i < $maxWeek; $i++) {
                if ($singleOrDouble === 1 && $i % 2 !== 0) continue;
                if ($singleOrDouble === 2 && $i % 2 === 0) continue;

                $date = Carbon::createFromDate($startDate[0], $startDate[1], $startDate[2], 'Asia/Shanghai');
                $date->addWeeks($i)->addDays($day);

                $isSummerTime = $date->isBetween(Carbon::createFromDate($startDate[0], 5, 1, 'Asia/Shanghai'), Carbon::createFromDate($startDate[0], 10, 1, 'Asia/Shanghai'));
                $data[] = $name;
                $data[] = $date->toDateString();
                $data[] = $date->toDateString();
                $data[] = getClassStartTime($startNum, $isSummerTime);
                $data[] = getClassEndTime($endNum, $isSummerTime);
                $data[] = $location;
                $data[] = "$teacherName $classNum";

                $csv_filed[] = $data;
                $data = array();
            }
        }
    }

    csvExport('课表', $csv_title, $csv_filed);
} catch (Exception $e) {
    return false;
}

