<?php

require './vendor/autoload.php';
require './helper.php';

use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\IOFactory;

try {
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

        $duration = explode(',', $worksheet->getCell("E$index")->getValue());
        foreach ($duration as $value) {
            $matches = array();
            preg_match_all('/\d+/', $value, $matches);
            @$maxWeek = $matches[0][1] !== 0 ? (int)$matches[0][1] : $matches[0][0];
            for ($i = $matches[0][0] - 1; $i < $maxWeek; $i++) {
                $date = Carbon::createFromDate(2019, 8, 26, 'Asia/Shanghai');
                $date->addWeeks($i)->addDays($day);

                $isSummerTime = $date->isBetween(Carbon::create(2019, 5, 1), Carbon::create(2019, 10, 1));
                $data[] = $name;
                $data[] = $date->toDateString();
                $data[] = $date->toDateString();
                $data[] = getClassStartTime($startNum, $isSummerTime);
                $data[] = getClassEndtTime($endNum, $isSummerTime);
                $data[] = $location;
                $data[] = "$teacherName $classNum";

                $csv_filed[] = $data;
                $data = array();
            }
        }
    }

    csvExport("课表", $csv_title , $csv_filed);
} catch (Exception $e) {
    return false;
}

