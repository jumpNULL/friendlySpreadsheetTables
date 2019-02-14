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
 *    Add support for aligning elements based on separators in cell values, header and footer.
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
    private $divisorFunction;
    private $defaultStyles = [
        'defaultTableHeaderStyle' => [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => [
                    'argb' => 'FFA0A0A0'
                ],
            ]
        ],
        'defaultTableFooterStyle' => [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
            ],
            'font' => [
                'bold' => true
            ]
        ],
        'defaultColumnHeaderStyle' => [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ]
        ],
        'defaultColumnFooterStyle' => [
            'font' => [
                'bold' => true
            ]
        ],
        'defaultTableStyle' => [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
                ],
                'inside' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
                ]
            ]
        ]
    ];

    public function __construct(Worksheet $sheet, $anchorColumn = 'B', $anchorRow = 2)
    {
        if (!Utility::validSheetCell($anchorColumn, $anchorRow)) {
            throw new InvalidSheetCellAddressException($anchorColumn . $anchorRow);
        }
        $this->sheet = $sheet;
        $this->anchorCell = (object)['column' => $anchorColumn, 'row' => $anchorRow];
        $this->tables = array();
    }

    public function export(): SpreadsheetTableFacade
    {
        $this->sheet->getDefaultRowDimension()->setRowHeight(15);
        $anchorCell = clone $this->anchorCell;

        foreach ($this->tables as $table) {

            $table->anchor($anchorCell->column, $anchorCell->row);
            $this->renderTable($table);
            //update point
            $lowerRightCell = $table->getLowerRightCell();
            $anchorCell->row = ($lowerRightCell->row++);

            //TODO: Separator logic
            if ($this->getDivisorFunction() !== NULL) {
                call_user_func($this->getDivisorFunction(), [$this->sheet, $this->anchorCell]);
            }
            else {
                $anchorCell->row += 2;
            }
        }
        //after all tables are anchored and rendered we resize all columns within their ranges
        foreach($this->tables as $table){
            for( $column = $table->getAnchor()->column, $tableEndColumn = $table->getLowerRightCell()->column; $column <= $tableEndColumn; $column++)
            {
                $this->sheet->getColumnDimension($column)->setAutoSize(true);
            }
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
        //apply style over range
        $this->sheet->getStyle($table->getAnchorAddress() . ':' . $table->getLowerRightCellAddress())->applyFromArray($table->getStyleArray());

        if ($table->getHeader() !== NULL) {
            if ($table->getSheetCellWidth() > 1) {
                try {
                    $table->getHeader()->setSheetCellWidth($table->getSheetCellWidth());
                } catch (\Exception $e) {
                    throw $e;
                };
            }
            $this->renderCell($table->getHeader());
        }
        //render sub-elements
        foreach ($table->getElements() as $entity) {
            //TODO: Footers and header elements for table entities should all be rendered at
            //    the same level to ensure visual consistency.
            if (is_a($entity, SheetTableColumn::class)) {
                $this->renderColumn($entity);
            } elseif (is_a($entity, SheetTable::class)) {
                $this->renderTable($entity);
            }
        }

        //Footer
        if ($table->getFooter() !== NULL) {
            if ($table->getSheetCellWidth() > 1) {
                try {
                    $table->getFooter()->setSheetCellWidth($table->getSheetCellWidth());
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
     * @return SpreadsheetTableFacade
     * @throws UnanchoredException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function renderColumn(SheetTableColumn $column): SpreadsheetTableFacade
    {

        $this->sheet->getStyle($column->getAnchorAddress() . ':' . $column->getLowerRightCellAddress())->applyFromArray($column->getStyleArray());
        //If column size is greater than unity, merge cells to which column will map
        foreach ($column->getCells() as $cell) {
            $this->renderCell($cell);
        }
        return $this;
    }

    /**
     * @param SheetTableCell $cell
     *
     * @return SpreadsheetTableFacade
     * @throws \Exception
     * @todo Add data-type conditional formatting logic (numbers right-aligned,
     *       test left-aligned, date and monetary localization)
     *
     */
    protected function renderCell(SheetTableCell $cell): SpreadsheetTableFacade
    {
        //Pre-condition: cell->anchor() was successfully called (getAnchorAddress() !== NULL)
        try {
            if ($cell->getSheetCellWidth() > 1 || $cell->getSheetCellHeight() > 1) {
                $this->sheet->mergeCells($cell->getAnchorAddress() . ':' . $cell->getLowerRightCellAddress());
            }
            $sheetCell = $this->sheet->getCell($cell->getAnchorAddress());
            if ($sheetCell !== NULL) {
                $sheetCell->setValue($cell->getValue());
            }
            if ($sheetCell !== NULL) {
                $sheetCell->getStyle()->applyFromArray($cell->getStyleArray());
            }
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
            if ($table !== NULL && is_object($table)) {
                $this->tables[] = $table;
            }
        }
        return $this;
    }

    /**
     * Divisor function should accept a spreadsheet parameter and current anchor.
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
     * @return callable|null
     */
    public function getDivisorFunction(): ?callable
    {
        return $this->divisorFunction;
    }

    /**
     * Applies default styling to tables and columns. Provided for convenience.
     *
     * @param SheetTable[] $tables
     *
     * @return SpreadsheetTableFacade
     */
    public function applyDefaultStyle(SheetTable ...$tables): self
    {
        foreach($tables as $table) {
            $table->setStyleArray($this->defaultStyles['defaultTableStyle']);
            $table->setHeaderStyleArray($this->defaultStyles['defaultTableHeaderStyle']);
            $table->setFooterStyleArray($this->defaultStyles['defaultTableFooterStyle']);
            foreach($table->getElements() as $element) {
                if(get_class($element) == SheetTableColumn::class) {
                    $element->setHeaderStyleArray($this->defaultStyles['defaultColumnHeaderStyle']);
                    $element->setFooterStyleArray($this->defaultStyles['defaultColumnFooterStyle']);
                }
            }
        }
        return $this;
    }
}

