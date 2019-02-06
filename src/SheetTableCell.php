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
class SheetTableCell {
    /**
     * @var array The upper left-hand sheet cell which will act as the origin
     *            for all sheet placement calculations with respect to the table
     *            cell.
     */
    private $cellAnchor;

    /**
     * @var array Array of PhpSpreadsheet styles to apply.
     *
     * @link https://phpspreadsheet.readthedocs.io/en/develop/topics/recipes/#styles
     */
    private $styleArray;

    /**
     * @var array Table cells dimensions are measured in sheet cells. A table cell
     *            may have non-unity width and height. In such a case, the sheet
     *            cells surrounding the anchor cell are merged to create the table
     *            cell.
     */
    private $cellDimensions;

    /**
     * @var \stdClass The cells value. If the cells value is not a number or float,
     *                then the value is cast to a string.
     */
    private $value;

    /**
     * SheetTableCell constructor.
     *
     * @param int $sheetCellWidth
     * @param int $sheetCellHeight
     *
     * @throws InvalidTableCellDimensionException
     */
    public function __construct(int $sheetCellWidth = 1, int $sheetCellHeight = 1)
    {
        $this->cellAnchor = array();
        $this->styleArray = array();

        if($sheetCellWidth <= 0) {
            throw new InvalidTableCellDimensionException('Width', $sheetCellWidth);
        } elseif ($sheetCellHeight <= 0){
            throw new InvalidTableCellDimensionException('Height', $sheetCellHeight);
        }

        $this->cellDimensions['cellWidth'] = $sheetCellWidth;
        $this->cellDimensions['cellHeight'] = $sheetCellHeight;
    }

    /**
     * Anchor a tables cell to a sheet cell. This is necessary for calculating
     * sheet range properties and a prerequisite to attaching to a sheet.
     *
     * @param string $cellColumn
     * @param int    $cellRow
     *
     * @return SheetTableCell
     * @throws InvalidSheetCellAddressException
     */
    public function anchor(string $cellColumn, int $cellRow) : SheetTableCell
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
     * @return \stdClass
     */
    public function getValue() : \stdClass
    {
        return $this->value;
    }


    /**
     * @return int
     */
    public function getSheetCellHeight() : int
    {
        return $this->cellDimensions['height'];
    }

    /**
     * @return int
     */
    public function getSheetCellWidth() : int
    {
        return $this->cellDimensions['width'];
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
        $this->value = $value;

        return $this;
    }

    /**
     * @param int $cellWidth
     *
     * @return SheetTableCell
     *
     * @throws InvalidTableCellDimensionException
     */
    public function setSheetCellWidth(int $cellWidth) : SheetTableCell
    {
        if($cellWidth <= 0)
            throw new InvalidTableCellDimensionException('Width', $cellWidth);
        else
            $this->cellDimensions['cellWidth'] = $cellWidth;

        return $this;
    }

    /**
     * @param int $cellHeight
     *
     * @return SheetTableCell
     *
     * @throws InvalidTableCellDimensionException
     */
    public function setSheetCellHeight(int $cellHeight) : SheetTableCell
    {
        if($cellHeight <= 0)
            throw new InvalidTableCellDimensionException('Height', $cellHeight);
        else
            $this->cellDimensions['cellHeight'] = $cellHeight;

        return $this;
    }

}