<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load("uploads/latestData.xlsx");

$worksheet = $spreadsheet->getActiveSheet();
// Get the highest row and column numbers referenced in the worksheet
$highestRow = $worksheet->getHighestDataRow()-1; // e.g. 10
$highestColumn = $worksheet->getHighestDataColumn(); // e.g 'F'
$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); // e.g. 5

// the first column should be +
$col = 1;
$check_failed=false;

//it should contain only one sheet
$sheetCount = $spreadsheet->getSheetCount();
if($sheetCount != 1) {
    $check_failed=true;
    $failed_message="This is not the latest data as it contains more than one worksheet.";
}
//Cells(1, 1).Value = "Date"
if(!$check_failed) {
    $value = $worksheet->getCell([1, 1])->getValue();
    if($value != 'Date') {
        $check_failed=true;
        $failed_message="This is not the latest data as it first column header should be Date";
    }
}

//Cells(1, 3).Value = "Company Name"
if(!$check_failed) {
    $value = $worksheet->getCell([3, 1])->getValue(); 
    if($value != 'Company Name') {
        $check_failed=true;
        $failed_message="This is not the latest data as it first column header should be Company Name. Value is ".$value;
    }
}

//the first column value should be +
if(!$check_failed) {
    for ($row = 2; $row <= $highestRow; ++$row) {
        $value = $worksheet->getCell([$col, $row])->getValue();
        if($value != '+') {
            $check_failed=true;
            $failed_message="This is not the latest data as it first column does not contain + value";
            //echo 'Failed at row '.$row;
        } else {
            //echo 'Good at row '.$row;
        }
    
    }
}


if($check_failed) {
    echo $failed_message . PHP_EOL;
} else {
    //echo 'So far so good.' . PHP_EOL;
    header("Location: uploadLatestReport.php");
    die();
}


echo '<table>' . "\n";
for ($row = 1; $row <= 4; ++$row) {
    echo '<tr>' . PHP_EOL;
    for ($col = 1; $col <= $highestColumnIndex; ++$col) {
        $value = $worksheet->getCell([$col, $row])->getValue();
        echo '<td>' . $value . '</td>' . PHP_EOL;
    }
    echo '</tr>' . PHP_EOL;
}
echo '</table>' . PHP_EOL;