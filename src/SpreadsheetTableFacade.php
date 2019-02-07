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
 *
 *
 * @link https://phpspreadsheet.readthedocs.io/en/develop/ PHPSpreadsheet Documentation
 *
 * @package Crombi\PhpSpreadsheetHelper
 */
class SpreadsheetTableFacade {

    private $tables;

    /**
     * @var
     */
    private $divisorFunction;

    public function __construct($table)
    {
        $this->tables = array();
    }

    /**
     * @return Spreadsheet
     */
    public function export () : Spreadsheet
    {
        $spreadsheet = new Spreadsheet();

        try {
            $sheet = $spreadsheet->getActiveSheet();
        } catch (\Exception $e) {
            return NULL;
        }



        return $spreadsheet;
    }

    /**
     * The facade allows rendering multiple tables to a sheet. Each
     * table is separated by a divider.
     *
     * @param SheetTable $table
     *
     * @return SpreadsheetTableFacade
     */
    public function addTables(SheetTable $table) : SpreadsheetTableFacade
    {
        if(!is_null($table) && is_object($table))
            array_push($this->tables, $table);

        return $this;
    }

    /**
     * Divisor function should accept a spreadsheet parameter.
     *
     * @param callable $divisorFunction
     *
     * @return SpreadsheetTableFacade
     */
    public function setDivisorFunction(callable $divisorFunction) : SpreadsheetTableFacade
    {

        return $this;
    }

    public function getDivisorFunction() : callable
    {
        return $this->divisorFunction;
    }
}
