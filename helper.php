<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;

function getClassStartTime($startNum, $isSummerTime) {
    switch ($startNum) {
        case 1:
            return '8:30';
        case 3:
            return '10:25';
        case 5:
            if ($isSummerTime) {
                return '14:30';
            }
            return '14:00';
        case 7:
            if ($isSummerTime) {
                return '16:25';
            }
            return '15:55';
        case 9:
            if ($isSummerTime) {
                return '19:30';
            }
            return '19:00';
    }
    return '';
}

function getClassEndTime($endNum, $isSummerTime) {
    switch ($endNum) {
        case 2:
            return '10:05';
        case 4:
            return '12:00';
        case 6:
            if ($isSummerTime) {
                return '16:05';
            }
            return '15:35';
        case 8:
            if ($isSummerTime) {
                return '18:00';
            }
            return '17:30';
        case 10:
            if ($isSummerTime) {
                return '21:05';
            }
            return '20:35';
    }
    return '';
}

function csvExport($excel_name = 'test', $excel_title = [], $excel_filed = []) {
    try {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        @ob_clean();
        $cellName = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];

        foreach ($excel_title as $key => $value) {
            $sheet->setCellValue("$cellName[$key]1", $value);
        }

        foreach ($excel_filed as $rowNum => $item) {
            foreach ($item as $columnNam => $value) {
                $sheet->setCellValue($cellName[$columnNam] . ($rowNum + 2), $value);
            }
        }

        header('Content-Type: text/csv');
        header('Content-Disposition: attachment;filename="' . $excel_name . '.csv"');
        header('Cache-Control: max-age=0');
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
        $writer = new Csv($spreadsheet);
        $writer->save('php://output');
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        exit;
    } catch (Exception $e) {
        return false;
    }
}