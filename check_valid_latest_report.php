<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load("uploads/latestReport.xlsx");


// the first column should be +
$col = 1;
$check_failed=false;

//it should contain only one sheet
$sheetCount = $spreadsheet->getSheetCount();
if($sheetCount == 1) {
    $check_failed=true;
    $failed_message="This is not the latest report as it contains only one worksheet.";
}

if(!$check_failed) {
    $worksheet = $spreadsheet->getSheet(1);
    // Get the highest row and column numbers referenced in the worksheet
    $highestRow = $worksheet->getHighestDataRow()-1; // e.g. 10
    $highestColumn = $worksheet->getHighestDataColumn(); // e.g 'F'
    $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); // e.g. 5
    
}

//Cells(1, 1).Value = "Date"
if(!$check_failed) {
    $value = $worksheet->getCell([1, 1])->getValue();
    if($value != 'Date') {
        $check_failed=true;
        $failed_message="This is not the latest report as it first column header should be Date";
    }
}

//Cells(1, 2).Value = "Company Name"
if(!$check_failed) {
    $value = $worksheet->getCell([2, 1])->getValue(); 
    if($value != 'Company Name') {
        $check_failed=true;
        $failed_message="This is not the latest report as it first column header should be Company Name. Value is ".$value;
    }
}

if($check_failed) {
    echo $failed_message . PHP_EOL;
} else {
    //echo 'So far so good. Next is to process the excel file.' . PHP_EOL;
    header("Location: process_excelfile.php");
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

?>
<form action="createReport.php" method="post" enctype="multipart/form-data">
  <input type="submit" value="Restart" name="submit">
</form>