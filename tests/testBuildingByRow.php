<?php
require '../vendor/autoload.php';
//The only reason theres so many () is because PHP won't let us method
//chain off a constructor if we don't wrap it in parentheses
use \Crombi\PhpSpreadsheetHelper\SpreadsheetTableFacade;
use \Crombi\PhpSpreadsheetHelper\SheetTableColumn;
use \Crombi\PhpSpreadsheetHelper\SheetTable;
use \PhpOffice\PhpSpreadsheet\Writer\Xls;

$spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$facade = new SpreadsheetTableFacade($spreadsheet->getActiveSheet());

$table = (new SheetTable())->addElements(
    new SheetTableColumn(),
    new SheetTableColumn(),
    new SheetTableColumn()
);

$table->addValues(
    [1],
    [5, 6],
    [7, 8, 9], //When data count is greater than column count, extra data is discarded and exception thrown
    [10, 11, 12, 13]
);

$facade->addTables($table)->export();
//Write out the PhpSpreadsheet using the PhpSpreadsheet writer
$writer = new Xls($spreadsheet);
$writer->save("test.xls");
