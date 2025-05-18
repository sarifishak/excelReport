<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$newReportSpreadsheet = new Spreadsheet();
$newReportActiveWorksheet = $newReportSpreadsheet->getActiveSheet();

//read from the latest data, first
$latestDataReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$latestDataReader->setReadDataOnly(TRUE);
$latestDataSpreadsheet = $latestDataReader->load("uploads/latestData.xlsx");

$latestDataWorksheet = $latestDataSpreadsheet->getActiveSheet();
// Get the highest row and column numbers referenced in the worksheet
$latestDataHighestRow = $latestDataWorksheet->getHighestDataRow()-1; // e.g. 10
$latestDataHighestColumn = $latestDataWorksheet->getHighestDataColumn(); // e.g 'F'
$latestDataHighestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($latestDataHighestColumn); // e.g. 5

for ($row = 1; $row <= $latestDataHighestRow; ++$row) {

    for ($col = 2; $col <= $latestDataHighestColumnIndex; ++$col) {
        $value = $latestDataWorksheet->getCell([$col, $row])->getValue();
        $newReportActiveWorksheet->setCellValue([$col-1, $row], $value);
    }

}

$newReportActiveWorksheet->setCellValue([1, 1], 'Date');

$newReportWriter = new Xlsx($newReportSpreadsheet);
$newReportWriter->save('uploads/newReport.xlsx');
?>
<H1>Report is created and sent to your personal email</H1>
<form action="createReport.php" method="post" enctype="multipart/form-data">
  <input type="submit" value="Restart" name="submit">
</form>