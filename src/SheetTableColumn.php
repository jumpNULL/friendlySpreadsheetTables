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
class SheetTableColumn
{
    private $cellAnchor;
    private $footerStyleArray;
    private $headerStyleArray;
    private $cellWidth;

    /**
     * @var bool Prevents the cell width count of a column being updated once a
     *           value has been added.
     */
    private $lockedWidth;
    private $title;
    private $footer;
    private $sheetCells;

    /**
     * SheetTableColumn constructor.
     *
     * @param array          $sheetCells
     * @param \stdClass|NULL $title
     * @param \stdClass|NULL $footer
     * @param array          $footerStyleArray
     * @param array          $headerStyleArray
     */
    public function __construct(array $sheetCells = array(), \stdClass $title = NULL,
                                \stdClass $footer = NULL, array $footerStyleArray = array(),
                                array $headerStyleArray = array())
    {
        //An empty string would be a cell with an empty title, NULL
        //is _no_ cell
        $this->title = $title;
        $this->footer = $footer;
        $this->sheetCells = $sheetCells;
        $this->lockedWidth = false;
        $this->cellAnchor = array();
        $this->footerStyleArray = $footerStyleArray;
        $this->headerStyleArray = $headerStyleArray;
    }

    //------------------------TABLE METHODS----------------------//
    /**
     * Anchor a tables cell to a sheet cell. This is necessary for calculating
     * sheet range properties and a prerequisite to attaching to a sheet.
     *
     * @param string $cellColumn
     * @param int    $cellRow
     *
     * @return SheetTableColumn
     * @throws InvalidSheetCellAddressException
     */
    public function anchor(string $cellColumn, int $cellRow) : SheetTableColumn
    {
        if (Utility::validSheetCell($cellColumn, $cellRow)){
            $this->cellAnchor['column'] = $cellColumn;
            $this->cellAnchor['row'] = $cellRow;
        } else
            throw new InvalidSheetCellAddressException($cellColumn .
                strval($cellRow));

        return $this;
    }

    /**
     * @param \stdClass $header
     *
     * @return SheetTableColumn
     */
    public function setHeader(\stdClass $header) : SheetTableColumn
    {
        return $this;
    }

    /**
     * @param \stdClass $footer
     *
     * @return SheetTableColumn
     */
    public function setFooter(\stdClass $footer) : SheetTableColumn
    {
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
                $sheetTableCell = new SheetTableCell($value);
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
        return array_map(function (SheetTableCell $ele) : \stdClass
        {
            return $ele->getValue();
        }, $this->sheetCells);
    }

    /**
     * @return int
     */
    public function getCellRangeWidth() : int
    {

    }

    /**
     * @return int
     */
    public function getCellRangeHeight() : int
    {

    }

    /**
     * @param array $styleArray
     *
     * @return SheetTableColumn
     */
    public function setTitleStyleArray (array $styleArray) : SheetTableColumn
    {
        return $this;
    }

    /**
     * @param array $styleArray
     *
     * @return SheetTableColumn
     */
    public function setFooterStyleArray(array $styleArray) : SheetTableColumn
    {
        return $this;
    }
}
