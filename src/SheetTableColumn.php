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
    private $upperLeftCell;
    private $footerStyleArray;
    private $titleStyleArray;

    private $cellWidth;

    /**
     * @var bool Prevents the cell width count of a column being updated once a
     *           value has been added.
     */
    private $lockedWidth;
    private $title;
    private $footer;
    private $sheetCells;

    public function __construct()
    {
        $this->title = NULL;
        $this->footer = NULL;
        $this->sheetCells = array();
    }

    //------------------------TABLE METHODS----------------------//
    public function setTitle($title)
    {
        return $this;
    }

    public function setFooter($footer)
    {
        return $this;
    }

    public function addValues($values)
    {
        foreach ($values as $value)
        {
            array_push($this->sheetCells, new Cell());
        }

        return $this;
    }

    public function getValues()
    {
        return array_map(function ($ele)
        {
            return $ele->getValue();
        }, $this->sheetCells);
    }
    //------------------SHEET RENDERING METHODS------------------//
    public function attach($cellColumn, $cellRow)
    {

    }

    public function getCellRangeWidth()
    {

    }

    public function getCellRangeHeight()
    {

    }

    public function setTitleStyleArray ($array)
    {
        return $this;
    }

    public function setFooterStyleArray($array)
    {
        return $this;
    }
}