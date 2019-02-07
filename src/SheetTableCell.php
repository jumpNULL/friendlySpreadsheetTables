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
class SheetTableCell implements TableEntityInterface
{
    use AnchorableTrait;

    /**
     * @var array Array of PhpSpreadsheet styles to apply.
     *
     * @link https://phpspreadsheet.readthedocs.io/en/develop/topics/recipes/#styles
     */
    private $styleArray;

    /**
     * @var \stdClass The cells value. If the cells value is not a number or float,
     *                then the value is cast to a string.
     */
    private $value;

    /**
     * SheetTableCell constructor.
     *
     * @param \stdClass $value
     * @param int $sheetCellWidth
     * @param int $sheetCellHeight
     * @param array $styleArray
     *
     * @throws InvalidTableCellDimensionException
     */
    public function __construct(\stdClass $value, int $sheetCellWidth = 1, int $sheetCellHeight = 1,
                                array $styleArray = array())
    {
        $this->styleArray = $styleArray;
        $this->setSheetCellHeight($sheetCellHeight);
        $this->setSheetCellWidth($sheetCellWidth);
        $this->setValue($value);
    }

    /**
     * @return \stdClass
     */
    public function getValue() : \stdClass
    {
        return $this->value;
    }

    /**
     * @return array
     */
    public function getStyleArray() : array
    {
        return $this->styleArray;
    }

    /**
     * @param \stdClass $value
     *
     * @return SheetTableCell
     */
    public function setValue(\stdClass $value) : SheetTableCell
    {
        if(!is_null($value))
            $this->value = $value;

        return $this;
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
}
