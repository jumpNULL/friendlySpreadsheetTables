<?php

use \PhpOffice\PhpSpreadsheet\Spreadsheet;
use \PhpOffice\PhpSpreadsheet\Writer\Xls;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

//Merge cells we will be using
//Merge Three
$sheet->mergeCells('A6:C6');
$sheet->mergeCells('D9:I9');
$sheet->mergeCells('D17:I17');
$sheet->mergeCells('D18:H18');
$sheet->mergeCells('F25:H25');

//Merge Two
for($row = 10; $row <= 15; $row++) {
    $sheet->mergeCells('D'.$row.':'.'F'.$row);
    $sheet->mergeCells('G'.$row.':'.'I'.$row);
}

for($row = 9; $row <= 23; $row++) {
    if($row == 16)
        continue;
    $sheet->mergeCells('A'.$row.':'.'C'.$row);
}

for($row = 25; $row <= 45; $row++) {
    if($row == 30)
        continue;
    $sheet->mergeCells('D'.$row.':'.'E'.$row);
}
//set values

//set styles
$titleStyle = [
    'alignment' => []
];

$sheet->getStyle('F25:H25')->applyFromArray();
$sheet->getStyle('D17:I17')->applyFromArray();
$sheet->getStyle('F31')->applyFromArray();
$sheet->getStyle('D9:I9')->applyFromArray();
$sheet->getStyle('D9:F9')->applyFromArray();
$sheet->getStyle('G9:I9')->applyFromArray();

$xls = new Xls($spreadsheet);
$xls->save('plainMovCaja');
