<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function findRecord($docNumber,$itemCode,$latestReportHighestRow,$latestReportWorksheet,$columnNoDocNo,$columnNoItemCode) {
    $recordNum = 0;
    for ($row = 1; $row <= $latestReportHighestRow; ++$row) {
        $docNo = $latestReportWorksheet->getCell([$columnNoDocNo, $row])->getValue();
        $itemC = $latestReportWorksheet->getCell([$columnNoItemCode, $row])->getValue();
        if($docNo === $docNumber && $itemC == $itemCode) {
            return $row;
        }
    }

    return $recordNum;
}

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

// copy first from the latestData
for ($row = 1; $row <= $latestDataHighestRow; ++$row) {

    for ($col = 2; $col <= $latestDataHighestColumnIndex; ++$col) {
        $value = $latestDataWorksheet->getCell([$col, $row])->getValue();
        $newReportActiveWorksheet->setCellValue([$col-1, $row], $value);
        // there is a problem of how to copy the style. Commented for now (18th May 2025)
        $styleArray = $latestDataWorksheet->getStyle([$col, $row])->exportArray();
        //echo "styleArray[".$col.",".$row."]:";
        //print_r($styleArray);
        $newReportActiveWorksheet->getStyle([$col-1, $row])->applyFromArray($styleArray);

    }

}

$newReportActiveWorksheet->setCellValue([1, 1], 'Date');

// next to get the data for column - "Shipment Status",	"VIA", "ETD", "ETA","DOCS","REMARK(S)"
$latestReportReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$latestReportReader->setReadDataOnly(TRUE);
$latestReportSpreadsheet = $latestReportReader->load("uploads/latestReport.xlsx");
$latestReportWorksheet = $latestReportSpreadsheet->getSheet(1);
// Get the highest row and column numbers referenced in the worksheet
$latestReportHighestRow = $latestReportWorksheet->getHighestDataRow()-1; // e.g. 10
$latestReportHighestColumn = $latestReportWorksheet->getHighestDataColumn(); // e.g 'F'
$latestReportHighestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($latestReportHighestColumn); // e.g. 5

//get the index of "Doc No","Item Code" and "Shipment Status" at row=1 
$row=1;
$columnNoDocNo=0;
$columnNoItemCode=0;
$columnNoShipmentStatus=0;
for ($col = 1; $col <= $latestReportHighestColumnIndex; ++$col) {
    $value = $latestReportWorksheet->getCell([$col, $row])->getValue();
    if($value === "Doc No") {
        $columnNoDocNo=$col;
    }
    if($value === "Item Code") {
        $columnNoItemCode=$col;
    }
    if($value === "Shipment Status") {
        $columnNoShipmentStatus=$col;
    }
}

//echo "<h2>columnNoDocNo is".$columnNoDocNo.",columnNoItemCode is ".$columnNoItemCode.",columnNoShipmentStatus is ".$columnNoShipmentStatus.",latestReportHighestColumnIndex=".$latestReportHighestColumnIndex."</h2>";

// Start copying the column header first for the newReportActiveWorksheet
if($columnNoShipmentStatus > 0) {
    $row=1;
    // Get the last column to be inserted
    $startColumn=1;
    for ($col = 1; $col <= $latestDataHighestColumnIndex; ++$col) {
        $value = $newReportActiveWorksheet->getCell([$col, $row])->getValue();
        if(strlen($value) == 0) {
            $startColumn = $col;
            break;
        }
    }
    $endColumn = $startColumn;
    for ($col = $columnNoShipmentStatus; $col <= $latestReportHighestColumnIndex; ++$col) {
        $value = $latestReportWorksheet->getCell([$col, $row])->getValue();
        //echo "value is ".$value.",col=".$col.",startColumn=",$startColumn;
        $newReportActiveWorksheet->setCellValue([$endColumn, $row], $value);
        $endColumn= $endColumn+1;
    }
    $newReportHighestColumnIndex =$endColumn;
}

// InsertTheRestOfData
for ($row = 1; $row <= $latestDataHighestRow; ++$row) {
    $docNumber = $newReportActiveWorksheet->getCell([3, $row])->getValue();
    $itemCode = $newReportActiveWorksheet->getCell([6, $row])->getValue();
    // check this record can be found from the latest report.
    $recordNum = findRecord($docNumber,$itemCode,$latestReportHighestRow,$latestReportWorksheet,$columnNoDocNo,$columnNoItemCode);
    if($recordNum != 0) {
        // echo "Record found for docnumber=".$docNumber.",itemcode=".$itemCode."<br>";
        $endColumn = $startColumn;
        for ($col = $columnNoShipmentStatus; $col <= $latestReportHighestColumnIndex; ++$col) {
            $value = $latestReportWorksheet->getCell([$col, $recordNum])->getValue();
            //echo "value is ".$value.",col=".$col.",startColumn=",$startColumn;
            $newReportActiveWorksheet->setCellValue([$endColumn, $row], $value);
            $endColumn= $endColumn+1;
        }
    } else {
        //echo "NO Record found for docnumber=".$docNumber.",itemcode=".$itemCode."<br>";
    }
}

// display the data in newReportActiveWorksheet
// echo '<table>' . "\n";
// for ($row = 1; $row <= $latestDataHighestRow; ++$row) {
//     echo '<tr>' . PHP_EOL;
//     for ($col = 1; $col <= $newReportHighestColumnIndex; ++$col) {
//         $value = $newReportActiveWorksheet->getCell([$col, $row])->getValue();
//         echo '<td>' . $value . '</td>' . PHP_EOL;
//     }
//     echo '</tr>' . PHP_EOL;
// }
// echo '</table>' . PHP_EOL;

$newReportWriter = new Xlsx($newReportSpreadsheet);
$newReportWriter->save('uploads/newReport.xlsx');

$filePath = 'uploads/newReport.xlsx'; 
$file = 'newReport.xlsx';
echo 'Please download this report <a href="' . $filePath . '" download>' . ucfirst(pathinfo($file, PATHINFO_FILENAME)) . '</a><br>'; 
?>
<form action="createReport.php" method="post" enctype="multipart/form-data">
  <input type="submit" value="Restart" name="submit">
</form>