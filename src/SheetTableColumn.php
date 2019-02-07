<?php declare(strict_types=1);
namespace Crombi\PhpSpreadsheetHelper;

use mysql_xdevapi\Exception;

/**
 * Class SheetTableColumn
 * An ordered set of values rendered vertically, each one below the last.
 *
 * A SheetTableColumn consist of an ordered set of cells. Optionally, it may
 * also possess a title and a footer. The title is rendered before all value
 * cells, and the footer after all value cells. Simply stated, a column may be
 * viewed as an atomic table. Styles may be applied to the footer, header or
 * collection of cells (but not cells individually).
 *
 * All cells in a column have the same width, but may have varying height. Thus
 * the width of each cell is the width of the column, and the height of the
 * column is the sum of the heights of each table cell which comprises it.
 *
 * @see SheetTableCell
 * @package Crombi\PhpSpreadsheetHelper
 */
class SheetTableColumn extends AnchorableEntity
{

    private $cellWidth;

    /**
     * @var bool Prevents the cell width of a column being updated once a
     *           value has been added.
     */
    private $lockedWidth;
    private $header;
    private $footer;
    private $sheetCells;
    private $styleArray;

    /**
     * SheetTableColumn constructor.
     *
     * @param array          $values
     * @param object|NULL $headerValue
     * @param object|NULL $footerValue
     * @param array          $footerStyleArray
     * @param array          $headerStyleArray
     */
    public function __construct(array $values = array(), object $headerValue = NULL,
                                array $headerStyleArray = array(), object $footerValue = NULL,
                                array $footerStyleArray = array() )
    {
        parent::__construct();
        //An empty string would be a cell with an empty title, NULL
        //is _no_ cell
        $this->sheetCells = array();
        $this->setHeader($headerValue);
        $this->setHeaderStyleArray($headerStyleArray);
        $this->setFooter($footerValue);
        $this->setFooterStyleArray($footerStyleArray);
        $this->addValues($values);
        $this->lockedWidth = false;

    }

    //------------------------TABLE METHODS----------------------//

    /**
     * Returns whether the tables column width is locked. This occurs after
     * adding values, and setting the header or  setting the footer.
     *
     * @return bool
     */
    public function isWidthLocked() : bool
    {
        return $this->lockedWidth;
    }

    //Updates anchor cell for all child table cells.
    public function resolveAddresses()
    {
        $cellColumn = $this->cellAnchor->column;
        $cellRow = $this->cellAnchor->row;

        //header cell comes first
        if(!is_null($this->header)) {
            $this->header->anchor($cellColumn, $cellRow);
            $cellRow += $this->header->getSheetCellWidth();
        }

        foreach ($this->sheetCells as $cell){
            $cell->anchor($cellColumn, $cellRow);
            $cellRow += $cell->getSheetCellWidth();
        }

        //footer cell comes last
        if(!is_null($this->footer))
            $this->footer->anchor($cellColumn, $cellRow);
    }

    /**
     * @param int $cellWidth
     *
     * @return SheetTableColumn
     * @throws InvalidTableCellDimensionException
     * @throws TableColumnWidthLocked
     */
    public function setSheetCellWidth(int $cellWidth) : SheetTableColumn
    {
        if(!$this->isWidthLocked())
            parent::setSheetCellWidth($cellWidth);
        else
            throw new TableColumnWidthLocked();
        return $this;
    }

    /**
     * @param object|NULL $headerValue Sets header to given value. If header value
     *                               is NULL, header is removed.
     *
     * @return SheetTableColumn
     */
    public function setHeader(?object $headerValue) : SheetTableColumn
    {
        $headerCell = NULL;

        try {
            if(!is_null($headerValue))
                $headerCell = new SheetTableCell($headerValue, $this->getSheetCellWidth());
        } catch (\Exception $e) {
            //Width exception should never occur, as if it does, it will occur
            //when setting the columns width
        }

        $this->header = $headerCell;

        return $this;
    }

    /**
     * @param object|NULL $footerValue Sets footer to given value. If header
     *                                    value is NULL, footer is removed.
     *
     * @return SheetTableColumn
     */
    public function setFooter(?object $footerValue) : SheetTableColumn
    {
        $footerCell = NULL;

        try {
            if(!is_null($footerValue))
                $footerCell = new SheetTableCell($footerValue, $this->getSheetCellWidth());
        } catch (\Exception $e) {
            //Width exception should never occur, as if it does, it will occur
            //when setting the columns width
        }

        $this->footer = $footerCell;

        return $this;
    }

    /**
     * @param array $values
     *
     * @return SheetTableColumn
     */
    public function addValues(array $values) : SheetTableColumn
    {
        foreach ($values as $value)
        {
            $this->lockedWidth = true;

            try {
                $sheetTableCell = new SheetTableCell((object) $value, $this->getSheetCellWidth());
            } catch (\Exception $e) {
                continue;
            }

            array_push($this->sheetCells, $sheetTableCell);
        }

        return $this;
    }

    /**
     * @return array
     */
    public function getValues() : array
    {
        return array_map(function (SheetTableCell $ele) : object
        {
            return $ele->getValue();
        }, $this->sheetCells);
    }

    /**
     * @param array $styleArray
     *
     * @return SheetTableColumn
     */
    public function setHeaderStyleArray (array $styleArray) : SheetTableColumn
    {
        if(!is_null($this->header))
            $this->header->setStyleArray($styleArray);
        return $this;
    }

    /**
     * @param array $styleArray
     *
     * @return SheetTableColumn
     */
    public function setFooterStyleArray(array $styleArray) : SheetTableColumn
    {
        if(!is_null($this->footer))
            $this->footer->setStyleArray($styleArray);

        return $this;
    }

    /**
     * Setting the height of a column is not supported. A column will have a
     * height equal to the sum of the heights of the cells that compose it.
     *
     * @param int $cellHeight
     *
     * @throws \Exception
     */
    public function setSheetCellHeight(int $cellHeight)
    {
        throw new \Exception('Invalid operation');
    }

    /**
     * The height of a column is equal to the sum of the heights of the cells
     * that compose it.
     *
     * @return int
     */
    public function getSheetCellHeight(): int
    {
        //Note that the height of a column is not equal to the number of cells multiplied
        //by a fixed height because each cell may have different size.
        return array_reduce($this->sheetCells, function (int $sum, SheetTableCell $cell) {
            return $sum + $cell->getSheetCellHeight();
    });
    }

    /**
     * Sets the columns style.
     *
     * @param array $styleArray
     *
     * @return SheetTableColumn
     */
    public function setStyleArray (array $styleArray) : SheetTableColumn
    {
        $this->styleArray = $styleArray;
        return $this;
    }

    public function getStyleArray () : array
    {
        return $this->styleArray;
    }

    /**
     * Applying a style to a column applies its style to every cell it contains.
     * It is important to note, however, that a cells style takes priority over
     * a columns style.
     *
     * @return SheetTableColumn
     */
    public function applyStyle() : SheetTableColumn
    {
        return $this;
    }

}
