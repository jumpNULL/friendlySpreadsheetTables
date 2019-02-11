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
 * @todo    Equations are currently not supported and must be added after building.
 *    Add support for alignment by separators in tuple values, header and footer.
 *
 * @todo    As a possible performance improvement point consider Settings::setCache()
 *       which expects a Psr compatible Cache implementation.
 *
 * @todo    Refactor entire render part.
 *
 * @link    https://phpspreadsheet.readthedocs.io/en/develop/ PHPSpreadsheet Documentation
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
    private $defaultHeaderStyle = [
        'alignment' => [
            'horizontal' =>  \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
        ],
        'font' => [
            'bold' => true
        ],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'color' => 'cecece'
        ]
    ];
    private $defaultFooterStyle = [];

    public function __construct(Worksheet $sheet, $anchorColumn = 'A', $anchorRow = 1)
    {
        if (!Utility::validSheetCell($anchorColumn, $anchorRow)) throw new InvalidSheetCellAddressException($anchorColumn . strval($anchorRow));
        $this->sheet = $sheet;
        $this->anchorCell = (object)['column' => $anchorColumn, 'row' => $anchorRow];
        $this->tables = array();
    }

    public function export() : SpreadsheetTableFacade
    {
        $firstTable = true;
        $anchorCell = clone $this->anchorCell;
        foreach ($this->tables as $table) {
            if (!$firstTable) {
                //TODO: Separator logic
                //render separator
            }
            $table->anchor($anchorCell->column, $anchorCell->row);
            $this->renderTable($table);
            //update point
            $lowerRightCell = $table->getLowerRightCell();
            $anchorCell->row = $lowerRightCell->row++;
            $firstTable=false;
        }

        return $this;
    }

    /**
     * @param SheetTable $table
     *
     * @return Worksheet
     * @throws InvalidSheetCellAddressException
     * @throws UnanchoredException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function renderTable(SheetTable $table): Worksheet
    {
        $anchorCell = $table->getAnchor();
        $rightCell = $table->getLowerRightCell();
        //render title
        if ($table->getHeader() !== NULL) {
            if ($table->getSheetCellWidth() > 1) {
                try {
                    $this->sheet->mergeCells($table->getAnchorAddress() . ':' . $rightCell->column . $anchorCell->row);
                } catch (\Exception $e) {
                    throw $e;
                };
            }
            $this->renderCell($table->getHeader());
        }

        //apply style over range
        $this->sheet->getStyle($table->getAnchorAddress() . ':' . $table->getLowerRightCellAddress())->applyFromArray($table->getStyleArray());
        //render sub-elements
        foreach ($table->getElements() as $entity) {
            if (is_a($entity, SheetTableColumn::class)) {
                $this->renderColumn($entity);
            } elseif (is_a($entity, SheetTable::class))
                $this->renderTable($entity);
        }

        if ($table->getFooter() !== NULL) {
            //render footer
            $rightCell = $table->getLowerRightCell();
            $rightCell->row = $table->getAnchor()->row;
            if ($table->getSheetCellWidth() > 1) {
                try {
                    $this->sheet->mergeCells($table->getAnchorAddress() . ':' . $rightCell->column . $rightCell->row);
                } catch (\Exception $e) {
                    throw $e;
                };
            }
            $this->renderCell($table->getFooter());
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
    public function renderColumn(SheetTableColumn $column): SpreadsheetTableFacade
    {
        if ($column->getSheetCellWidth() > 1) {
            for ($currentRow = $column->getAnchor()->row,
                 $anchorColumn = $column->getAnchor()->column,
                 $rightMostCell = $column->getLowerRightCell();
                 $currentRow <=  $rightMostCell->row; $currentRow++) {
                try {
                    $this->sheet->mergeCells($anchorColumn . $currentRow .
                    ':' . $rightMostCell->column . $currentRow);
                    //$this->sheet->getStyle($anchorColumn.$currentRow)->applyFromArray($column->getHeader()-);
                } catch (\Exception $e) {
                    throw $e;
                };
            }
        }

        $this->sheet->getStyle($column->getAnchorAddress() . ':' . $column->getLowerRightCellAddress());
        foreach ($column->getCells() as $cell) {
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
    protected function renderCell(SheetTableCell $cell): SpreadsheetTableFacade
    {
        //Pre-condition: cell->anchor() was successfully called (getAnchorAddress() !== NULL)
        //               $sheet->merge() on range was already called
        try {
            $sheetCell = $this->sheet->getCell($cell->getAnchorAddress());
            $sheetCell->setValue($cell->getValue());
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
    public function addTables(SheetTable ...$tables): SpreadsheetTableFacade
    {
        foreach ($tables as $table) {
            if (!is_null($table) && is_object($table)) array_push($this->tables, $table);
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
    public function setDivisorFunction(callable $divisorFunction): SpreadsheetTableFacade
    {
        $this->divisorFunction = $divisorFunction;
        return $this;
    }

    /**
     * @return callable
     */
    public function getDivisorFunction(): callable
    {
        return $this->divisorFunction;
    }
}

