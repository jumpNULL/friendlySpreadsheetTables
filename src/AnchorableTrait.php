<?php
namespace Crombi\PhpSpreadsheetHelper;

trait AnchorableTrait
{
    /**
     * @var array The upper left-hand sheet cell which will act as the origin
     *            for all sheet placement calculations with respect to the table
     *            cell.
     */
    private $cellAnchor = (object) array (
        'column' => NULL,
        'row' => NULL
    );

    /**
     * @var array Dimensions are measured in sheet cells. Given non-unity width
     *            and height, sheet cells surrounding the anchor cell are merged
     *            to create a single cell.
     */
    private $cellDimensions = (object) array (
        'width' => 1,
        'height' => 1
    );

    /**
     * @return int
     */
    public function getSheetCellHeight() : int
    {
        return $this->cellDimensions->height;
    }

    /**
     * @return int
     */
    public function getSheetCellWidth() : int
    {
        return $this->cellDimensions->width;
    }

    /**
     * Returns the address lower right cell.
     *
     * @return array
     *
     * @throws UnanchoredException
     */
    public function getLowerRightCell() : \stdClass
    {
        if(!$this->isAnchored())
            throw new UnanchoredException();

        if ($this->cellDimensions->width == 1 &&
            $this->cellDimensions->height == 1)
            return $this->cellAnchor;
        else {
            return (object) array(
                'column' => $this->cellAnchor->column + (--$this->cellDimensions->width),
                'row' => $this->cellAnchor->row + (--$this->cellDimensions->height)
            );
        }
    }

    /**
     * @param int $cellWidth
     *
     * @return TableEntityInterface
     *
     * @throws InvalidTableCellDimensionException
     */
    public function setSheetCellWidth(int $cellWidth) : TableEntityInterface
    {
        if($cellWidth <= 0)
            throw new InvalidTableCellDimensionException('Width', $cellWidth);
        else
            $this->cellDimensions->width = $cellWidth;

        return $this;
    }

    /**
     * @param int $cellHeight
     *
     * @return TableEntityInterface
     *
     * @throws InvalidTableCellDimensionException
     */
    public function setSheetCellHeight(int $cellHeight) : TableEntityInterface
    {
        if($cellHeight <= 0)
            throw new InvalidTableCellDimensionException('Height', $cellHeight);
        else
            $this->cellDimensions->height = $cellHeight;

        return $this;
    }

    /**
     * Anchor a tables element to a sheet cell. This is necessary for calculating
     * sheet range properties.
     *
     * @param string $cellColumn
     * @param int    $cellRow
     *
     * @return TableEntityInterface
     * @throws InvalidSheetCellAddressException
     */
    public function anchor(string $cellColumn, int $cellRow) : TableEntityInterface
    {
        if (Utility::validSheetCell($cellColumn, $cellRow)){
            $this->cellAnchor->column = $cellColumn;
            $this->cellAnchor->row    = $cellRow;
        } else
            throw new InvalidSheetCellAddressException($cellColumn .
                strval($cellRow));

        return $this;
    }

    /**
     * Returns whether a table element has been anchored to a sheet cell.
     *
     * @return bool
     */
    public function isAnchored() : bool
    {
        if (is_null($this->cellAnchor->width) || is_null($this->cellAnchor->height))
            return false;
        else
            return true;
    }
}
