<?php
require '../vendor/autoload.php';

/**
 * For now make sure that _ALL_ changes to table elements are done __before__
 * calling export() or anchor(), as changing width afterwards will not trigger
 * anchor updates causing layout to break. The alternative is to call
 * anchor() after any dimension changes to ensure model is maintained.
 */
use \Crombi\PhpSpreadsheetHelper\SheetTable;
use \Crombi\PhpSpreadsheetHelper\SheetTableColumn;
use \Crombi\PhpSpreadsheetHelper\SpreadsheetTableFacade;
use \PhpOffice\PhpSpreadsheet\Spreadsheet;
use \PhpOffice\PhpSpreadsheet\Writer\Xls;

$spreadsheet = new Spreadsheet();
$facade = new SpreadsheetTableFacade($spreadsheet->getActiveSheet());
$facade->applyDefaultStyle(true);

//Declare our tables
$tableArgentina = new SheetTable();
$tableForeign = new SheetTable();
$tableExpenditure = new SheetTable();
$tableOtherExpenses = new SheetTable();

//Build tables
$tableForeign->addElements(
    (new SheetTableColumn())->setSheetCellWidth(3)->addValues('', 'Ingresos', 'Interes cuenta', 'Interes Inversion', 'Dividendos / Utilidades'),
    (new SheetTable())->setHeader('FUENTE EXTRANJERA')->addElements(
        (new SheetTableColumn())->setSheetCellWidth(3)->setHeader('Gravado Tabla')->setFooter('Total'),
        (new SheetTableColumn())->setSheetCellWidth(3)->setHeader('No gravado/Convenio')->setFooter('Total')
    )
);

$tableArgentina->addElements(
    (new SheetTableColumn())->setSheetCellWidth(3)->addValues('', 'Ingresos', '', 'Interes inversion', 'Dividendos / Utilidades'),
    (new SheetTable())->setHeader('FUENTE ARGENTINA')->addElements(
        (new SheetTable())->setHeader('Gravado')->addElements(
            (new SheetTableColumn())->setHeader('5%')->setFooter('-'),
            (new SheetTableColumn())->setHeader('7%')->setFooter('-'),
            (new SheetTableColumn())->setHeader('13%')->setFooter('-'),
            (new SheetTableColumn())->setHeader('15%')->setFooter('-'),
            (new SheetTableColumn())->setHeader('Tabla')->setFooter('-')
    ),
        (new SheetTableColumn())->setHeader('No computable')
    )
);

$tableExpenditure->addElements(
    (new SheetTableColumn())->setSheetCellWidth(2)->addValues('', '',  'Interes', 'Gastos Varios', 'Gasto mantenimiento cuenta'),
    (new SheetTable())->setHeader('Gastos')->addElements(
        (new SheetTableColumn())->setHeader('Deducible'),
        (new SheetTableColumn())->setHeader('No'),
        (new SheetTableColumn())->setHeader('Sujeto a Prorrateo')
    )
);
$tableOtherExpenses->addElements(
    (new SheetTableColumn())->setSheetCellWidth(2)->setHeader('Otra informacion de interes')->addValues('Compras + Gs de compra',
        'Ventas - Gs de Venta', 'Amortizacion de Titulos', 'Rescates - gastos de rescate', 'Constitucion plazo fijo',
        'Cobranza plazo fijo (capital + interes)', 'Transferencias (-)', 'Depositos (+)', 'Retiros (-)', 'Pago de tarjeta de Credito',
        'Retencion impuesto a las ganancias', 'Impuesto analogo exterior', 'Diferencias de cambio'),
    (new SheetTableColumn())->setHeader('Total')
);

//Add them to facade for rendering
$facade->addTables($tableForeign, $tableArgentina, $tableExpenditure, $tableOtherExpenses)->export();

//Write out the PhpSpreadsheet using the PhpSpreadsheet writer
$writer = new Xls($spreadsheet);
$writer->save("resumen_caja.xls");
