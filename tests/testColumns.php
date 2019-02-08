<?php
require '../vendor/autoload.php';
//The only reason theres so many () is because PHP won't let us method
//chain off a constructor if we don't wrap it in parentheses
use \Crombi\PhpSpreadsheetHelper\SpreadsheetTableFacade;
use \Crombi\PhpSpreadsheetHelper\SheetTableColumn;
use \PhpOffice\PhpSpreadsheet\Writer\Xls;

$spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$facade = new SpreadsheetTableFacade($spreadsheet->getActiveSheet());
$column = new SheetTableColumn();

$column->addValues(['1', 2, 3])->setHeader('Test header')->setFooter('Test Footer');
$facade->renderColumn($column);

$writer = new Xls($spreadsheet);
$writer->save("test.xls");
