<?php declare(strict_types=1);
namespace Crombi\PhpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Class SheetTableBuilder
 * Provides logic for exporting the SpreadsheetHelper table model to a spreadsheet
 * utilizing PhpSpreadsheet.
 * **NOTE**: This class is _tightly_ coupled to the PHPSpreadsheet library.
 *
 * @todo Equations are currently not supported and must be added after building.
 *    Add support for alignment by separators in tuple values, header and footer.
 *    Decouple class from PHPSpreadsheet if possible to other providers.
 *
 * @link https://phpspreadsheet.readthedocs.io/en/develop/ PHPSpreadsheet Documentation
 *
 * @package Crombi\PhpSpreadsheetHelper
 */
class SpreadsheetTableFacade {

    public function __construct($table)
    {
    }

    public function export () : Spreadsheet
    {
        $spreadsheet = new Spreadsheet();

        return $spreadsheet;
    }
}
