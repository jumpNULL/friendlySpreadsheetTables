<?php
require '../vendor/autoload.php';
//The only reason theres so many () is because PHP won't let us method
//chain off a constructor if we don't wrap it in parentheses
use \Crombi\PhpSpreadsheetHelper\SpreadsheetTableFacade;
use \Crombi\PhpSpreadsheetHelper\SheetTable;
use \Crombi\PhpSpreadsheetHelper\SheetTableColumn;
use \PhpOffice\PhpSpreadsheet\Writer\Xls;

//Create SheetTable Model
$table = new SheetTable();

//Test variable width columns with mismatched data amount and building table out
//from the pre-populated columns
$table->addColumns([
    (new SheetTableColumn())->setTitle('ColumnOne')->setCellWidth(2)->addValues([1,2]),
    (new SheetTableColumn())->setTitle('ColumnTwo')->addValues([3])
]);

//Test adding data to table directly where data count is less than number of
//columns, equal to number of columns, and greater than number of columns
$table->addValues([
    [4], //When there is insufficient data
    [5, 6],
    [7, 8, 9] //When data count is greater than column count, extra data is discarded and exception thrown
]);

//Use the Facade to create a PhpSpreadsheet
$spreadsheet = (new SpreadsheetTableFacade($table))->export();

//Write out the PhpSpreadsheet using the PhpSpreadsheet writer
$writer = new Xls($spreadsheet);
$writer->save("test.xls");
