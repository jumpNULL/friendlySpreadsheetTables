<?php declare(strict_types=1);
namespace Crombi\PhpSpreadsheetHelper;

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
     * @param array       $values
     * @param int         $width
     * @param object|NULL $footerValue
     *
     * @throws InvalidTableCellDimensionException
     * @throws TableColumnWidthLocked
     */
    public function __construct(array $values = NULL, int $width = -1)
    {
        parent::__construct();

        $this->lockedWidth = false;
        $this->sheetCells = array();
        $this->setStyleArray([]);
        //We need a width if we are going to be adding in headers and footers
        if($width > 0)
            $this->setSheetCellWidth($width);
        else
            $this->setSheetCellWidth(1);
        //An empty string would be a cell with an empty title, NULL
        //is _no_ cell
        if(!is_null($values))
            $this->addValues($values);

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

    /**
     * @inheritdoc
     * @return mixed|void
     */
    protected function resolveAddresses()
    {
        $cellColumn = $this->cellAnchor->column;
        $cellRow = $this->cellAnchor->row;

        //header cell comes first
        if(!is_null($this->header)) {
            $this->header->anchor($cellColumn, $cellRow);
            $cellRow += 1;
        }

        foreach ($this->sheetCells as $cell){
            $cell->anchor($cellColumn, $cellRow);
            $cellRow += 1;
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
    public function setHeader($headerValue) : SheetTableColumn
    {
        $headerCell = $this->header;

        try {
            if(!is_null($headerValue)) {
                $headerCell = new SheetTableCell($headerValue, $this->getSheetCellWidth());
                $this->lockWidth();
            }
        } catch (\Exception $e) {
            //Width exception should never occur, as if it does, it will occur
            //when setting the columns width
        }

        $this->header = $headerCell;

        return $this;
    }

    /**
     * @param $footerValue Sets footer to given value. If header
     *                     value is NULL, footer is removed.
     *
     * @return SheetTableColumn
     */
    public function setFooter($footerValue) : SheetTableColumn
    {
        $footerCell = $this->footer;

        try {
            if(!is_null($footerValue)) {
                $footerCell = new SheetTableCell($footerValue, $this->getSheetCellWidth());
                $this->lockWidth();
            }
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
    public function addValues(...$values) : SheetTableColumn
    {
        foreach ($values as $value)
        {
            $this->lockWidth();

            try {
                $sheetTableCell = new SheetTableCell($value, $this->getSheetCellWidth());
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
        $cells = $this->getCells();
        //Note that the height of a column is not equal to the number of cells multiplied
        //by a fixed height because each cell may have different size.
        return array_reduce($cells, function (int $sum, SheetTableCell $cell) {
            return $sum + $cell->getSheetCellHeight();
        }, 0);
    }

    /**
     * Locks the columns width, preventing it from being altered further.
     */
    public function lockWidth() : void
    {
        $this->lockedWidth = true;
    }

    /**
     * Returns the array of cells associated with this column.
     *
     * ___WARNING__: DO NOT RELY ON THIS CLASS EXPOSING ITS UNDERLYING TableSheetCells.
     *
     * @todo This is a patch added on the fly. Exposing TableSheetCells is leaky
     *       design and a bad idea, as it opens up a whole host of edge cases
     *       where cell anchors and dimensions are edited _after_ being added
     *       to the container.
     *
     * @return array|null
     */
    public function getCells() : ?array
    {
        $cellArray = $this->sheetCells;

        if (!is_null($this->header)){
            $cellArray = array_merge([$this->header], $cellArray);
        }

        if(!is_null($this->footer)){
            $cellArray = array_merge($cellArray, [$this->footer]);
        }

        return $cellArray;
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

}
