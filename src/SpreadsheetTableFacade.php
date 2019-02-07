<?php declare(strict_types=1);
namespace Crombi\PhpSpreadsheetHelper;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Class SheetTableBuilder
 * Provides logic for exporting the SpreadsheetHelper table model to a spreadsheet
 * utilizing PhpSpreadsheet. Rendering is done by specifying an anchor for each table,
 * resolving address relative to the anchor, getting all table cells afterwards
 * and mapping them to their appropriate spreadsheet cell
 *
 * **NOTE**: This class is _tightly_ coupled to the PHPSpreadsheet library.
 *
 * @todo Equations are currently not supported and must be added after building.
 *    Add support for alignment by separators in tuple values, header and footer.
 *
 * @todo As a possible performance improvement point consider Settings::setCache()
 *       which expects a Psr compatible Cache implementation.
 * @link https://phpspreadsheet.readthedocs.io/en/develop/ PHPSpreadsheet Documentation
 *
 * @package Crombi\PhpSpreadsheetHelper
 */
class SpreadsheetTableFacade
{

    private $tables;
    private $anchorCell;
    private $tableType;
    /**
     * @var
     */
    private $divisorFunction;

    public function __construct(array $tables, $anchorColumn = 'A', $anchorRow = 1)
    {
        if(!Utility::validSheetCell($anchorColumn, $anchorRow))
            throw new InvalidSheetCellAddressException($anchorColumn .
                strval($anchorRow));

        $this->anchorCell = (object) [
          "column" => $anchorColumn,
          "row" => $anchorRow
        ];

        $this->setDivisorFunction(function(Worksheet $sheet ) {
            return;
        });

        $this->addTables($tables);
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

        //foreach table
            //Anchor table
            //resolve addresses
            //get list of table cells
            //for each cell
                //call ->merge('RC:R_C_') if cell dimensions are greater than unity
                //call ->getCell()
                //call the necessary ->getStyle()->applyFromArray()
                    //apply table, then column, then cell
                //apply datatype conditional logic (text is RichText, Dates need Locales,
                //money formatting, and numbers right aligned
                //call ->setValueValue() (possible alternative ->fromArray(values, ignore, topLeftCell)
            //render divisor
            //update anchor

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
        $this->divisorFunction = $divisorFunction;

        return $this;
    }

    public function getDivisorFunction() : callable
    {
        return $this->divisorFunction;
    }
}

