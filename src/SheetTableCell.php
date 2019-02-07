<?php
namespace Crombi\PhpSpreadsheetHelper;
/**
 * Class SheetCell
 * A entry associated with a style and size.
 *
 * Each cell contains a value, dimension and style. The dimension of a table
 * cell is measured in the number of surrounding cells which are assigned to it
 * and _not_ in terms of the size of a given sheet cell. The sheet cells which
 * are merged to create the table cell are calculated with origin of the anchor
 * cell.
 *
 * Contrary to the Table and Column model, cells do not posses headers or footers.
 *
 * @package Crombi\PhpSpreadsheetHelper
 */
class SheetTableCell extends AnchorableEntity
{

    /**
     * @var object The cells value. If the cells value is not a number or float,
     *                then the value is cast to a string.
     */
    private $value;
    private $styleArray;
    /**
     * SheetTableCell constructor.
     *
     * @param object $value
     * @param int $sheetCellWidth
     * @param int $sheetCellHeight
     * @param array $styleArray
     *
     * @throws InvalidTableCellDimensionException
     */
    public function __construct(object $value, int $sheetCellWidth = 1, int $sheetCellHeight = 1,
                                array $styleArray = array())
    {
        parent::__construct();
        $this->styleArray = $styleArray;
        $this->setSheetCellHeight($sheetCellHeight);
        $this->setSheetCellWidth($sheetCellWidth);
        $this->setValue($value);
    }

    /**
     * @return object
     */
    public function getValue() : object
    {
        return $this->value;
    }

    /**
     * @param object $value
     *
     * @return SheetTableCell
     */
    public function setValue(object $value) : SheetTableCell
    {
        if(!is_null($value))
            $this->value = $value;

        return $this;
    }

    /**
     * @return array
     */
    public function getStyleArray() : array
    {
        return $this->styleArray;
    }

    /**
     * @param array $styleArray
     *
     * @return SheetTableCell
     */
    public function setStyleArray(array $styleArray) : SheetTableCell
    {
        $this->styleArray = $styleArray;

        return $this;
    }

    /**
     * Applies a style to a cell.
     * @return SheetTableCell
     */
    public function applyStyle() : SheetTableCell
    {
        return $this;
    }

    /**
     * Table cells are atomic and have no post-anchor update logic. This
     * function simply returns the cell.
     * @return AnchorableEntity
     */
    public function resolveAddresses(): AnchorableEntity
    {
        return $this;
    }
}
