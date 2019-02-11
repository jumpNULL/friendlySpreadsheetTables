<?php
require '../vendor/autoload.php';
//The only reason theres so many () is because PHP won't let us method
//chain off a constructor if we don't wrap it in parentheses
use \Crombi\PhpSpreadsheetHelper\SpreadsheetTableFacade;
use \Crombi\PhpSpreadsheetHelper\SheetTableColumn;
use \PhpOffice\PhpSpreadsheet\Writer\Xls;

$spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$facade = new SpreadsheetTableFacade($spreadsheet->getActiveSheet());
$column = new SheetTableColumn(NULL, 2);
$columnTwo = new SheetTableColumn();

$table = new \Crombi\PhpSpreadsheetHelper\SheetTable();
$table->addElements($column, $columnTwo);

$column->addValues(...[1, 2])->setHeader('ColumnOne')->setFooter(3+4);
$columnTwo->addValues(...[3, 4])->setHeader('ColumnTwo')->setFooter('TestFooter');

$table->setHeader('Test Table');

$table->anchor('A', 1);

$facade->renderTable($table);

$writer = new Xls($spreadsheet);
$writer->save("test.xls");
