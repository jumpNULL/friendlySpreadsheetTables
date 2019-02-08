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
 *
 * @todo Refactor entire render part.
 *
 * @link https://phpspreadsheet.readthedocs.io/en/develop/ PHPSpreadsheet Documentation
 *
 * @package Crombi\PhpSpreadsheetHelper
 */
class SpreadsheetTableFacade
{

    private $tables;
    private $anchorCell;
    private $sheet;
    /**
     * @var
     */
    private $divisorFunction;

    public function __construct(Worksheet $sheet, array $tables = array(), $anchorColumn = 'A', $anchorRow = 1)
    {
        if(!Utility::validSheetCell($anchorColumn, $anchorRow))
            throw new InvalidSheetCellAddressException($anchorColumn .
                strval($anchorRow));

        $this->sheet =  $sheet;

        $this->anchorCell = (object) [
          'column' => $anchorColumn,
          'row' => $anchorRow
        ];

        $this->tables = array();
        $this->addTables($tables);
    }

    /**
     * @return Worksheet
     */
    public function export () : Worksheet
    {
        $firstTable = true;
        $anchorCell = clone $this->anchorCell;

        foreach($this->tables as $table) {
            if(!$firstTable) {
                //render separator
            }

            $table->anchor($anchorCell->column, $anchorCell->row);

            //render title
            $rightCell = $table->getLowerRightCell();
            $rightCell->row = $anchorCell->row;

            if ($table->getSheetCellWidth() > 1) {
                try {
                   $this->sheet->mergeCells($table->getAnchorAddress()
                        . ':' . $rightCell->column . $rightCell->row);
                } catch (\Exception $e) {
                    throw $e;
                };
            }

            $this->renderCell($table->getHeader());

            //apply style over range
            $this->sheet->getStyle()->applyFromArray($table->getStyleArray());

            //render sub-elements
            foreach($table->getElements() as $entity){
                var_dump($entity);
                die();
                switch(gettype($entity)){
                    case 'l':
                        renderColumn();
                        break;
                    case 'a':
                        renderTable();
                        break;
                    default:
                        break;
                }
            }

            //render footer
            $rightCell = $table->getLowerRightCell();
            $rightCell->row = $anchorCell->row;

            if ($table->getSheetCellWidth() > 1) {
                try {
                    $this->sheet->mergeCells($table->getAnchorAddress()
                        . ':' . $rightCell->column . $rightCell->row);
                } catch (\Exception $e) {
                    throw $e;
                };
            }
            $this->renderCell($table->getFooter());

            //update point
            $lowerRightCell = $table->getLowerRightCell();
            $anchorCell->row = $lowerRightCell->row++;
        }

        return $this->sheet;
    }

    /**
     * @param SheetTableColumn $column
     *
     * @return Worksheet
     * @throws InvalidSheetCellAddressException
     * @throws UnanchoredException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function renderColumn (SheetTableColumn $column) : SpreadsheetTableFacade
    {

        $column->anchor($this->anchorCell->column, $this->anchorCell->row);

        if ($column->getSheetCellWidth() > 1) {
            try {
                $this->sheet->mergeCells($column->getAnchorAddress()
                    . ':' . $column->getLowerRightCellAddress());
            } catch (\Exception $e) {
                throw $e;
            };
        }

            $this->sheet->getStyle($column->getAnchorAddress()
                . ':' . $column->getLowerRightCellAddress());

        foreach($column->getCells() as $cell) {
                $this->renderCell($cell);
        }

        return $this;
    }

    /**
     * @param SheetTableCell $cell
     *
     * @throws \Exception
     *
     * @todo Add data-type conditional formatting logic (numbers right-aligned,
     *       test left-aligned, date and monetary localization)
     *
     * @return Worksheet
     */
    protected function renderCell (SheetTableCell $cell) : SpreadsheetTableFacade
    {
        //Pre-condition: cell->anchor() was successfully called (getAnchorAddress() !== NULL)
        //               $sheet->merge() on range was already called
        try {
            $sheetCell = $this->sheet->getCell($cell->getAnchorAddress());
            $value = $cell->getValue();
            var_dump($value);
            $sheetCell->setValue($value);
            $sheetCell->getStyle()->applyFromArray($cell->getStyleArray());
        } catch (\Exception $e) {
            throw $e;
       }

        return $this;
    }

    /**
     * The facade allows rendering multiple tables to a sheet. Each
     * table is separated by a divider.
     *
     * @param SheetTable[] $tables
     *
     * @return SpreadsheetTableFacade
     */
    public function addTables(array $tables) : SpreadsheetTableFacade
    {
        foreach($tables as $table) {
            if (!is_null($table) && is_object($table))
                array_push($this->tables, $table);
        }

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

    /**
     * @return callable
     */
    public function getDivisorFunction() : callable
    {
        return $this->divisorFunction;
    }
}

