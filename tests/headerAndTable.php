<?php
require '../vendor/autoload.php';

use \PhpOffice\PhpSpreadsheet\Spreadsheet;
use \PhpOffice\PhpSpreadsheet\Writer\Xls;
use \Crombi\PhpSpreadsheetHelper\SpreadsheetTableFacade;
use \Crombi\PhpSpreadsheetHelper\SheetTable;
use \Crombi\PhpSpreadsheetHelper\SheetTableColumn;

$spreadsheet = new Spreadsheet();
$facade = new SpreadsheetTableFacade($spreadsheet->getActiveSheet());

$header = (new SheetTable())->addElements(
    new SheetTableColumn(NULL, 6),
    new SheetTableColumn(NULL, 6),
    new SheetTableColumn(NULL, 3)
);

$header->addValues(
    ['Período fiscal: 23/04/2018', 'Confeccionó:', 'Fecha:'],
    ['Contribuyente: John Jane-Doe', 'Revisó:', 'Fecha:']
);

$facade->applyDefaultStyle(true)->addTables($header)->export();

$writer = new Xls($spreadsheet);
$writer->save('header.xls');
