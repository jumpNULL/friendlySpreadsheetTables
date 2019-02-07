<?php
namespace Crombi\PhpSpreadsheetHelper;

abstract class AnchorableEntity
{
    /**
     * @var array The upper left-hand sheet cell which will act as the origin
     *            for all sheet placement calculations with respect to the table
     *            cell.
     */
    protected $cellAnchor;

    /**
     * @var array Dimensions are measured in sheet cells. Given non-unity width
     *            and height, sheet cells surrounding the anchor cell are merged
     *            to create a single cell.
     */
    private $cellDimensions;

    protected function __construct() {
        $this->cellAnchor = (object) [
            'column' => NULL,
            'row' => NULL
        ];

        $this->cellDimensions = (object) [
            'width' => 1,
            'height' => 1
        ];
    }

    /**
     * Extension point for allowing AnchorableEntity implementations
     * to update the anchor points of any children identities, or any other
     * post-anchor update conditional logic.
     *
     * @return mixed
     */

    abstract protected function resolveAddresses();
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
     * @return object
     *
     * @throws UnanchoredException
     */
    public function getLowerRightCell() : object
    {
        if(!$this->isAnchored())
            throw new UnanchoredException();

        if ($this->getSheetCellWidth() == 1 &&
            $this->getSheetCellHeight() == 1)
            return (object) $this->cellAnchor;
        else {
            $width = $this->getSheetCellWidth();
            $height = $this->getSheetCellHeight();

            return (object) array(
                'column' => $this->cellAnchor->column + --$width,
                'row' => $this->cellAnchor->row + --$height
            );
        }
    }

    /**
     * @param int $cellWidth
     *
     * @return AnchorableEntity
     *
     * @throws InvalidTableCellDimensionException
     */
    public function setSheetCellWidth(int $cellWidth)
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
     * @return AnchorableEntity
     *
     * @throws InvalidTableCellDimensionException
     */
    public function setSheetCellHeight(int $cellHeight)
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
     * @return AnchorableEntity
     * @throws InvalidSheetCellAddressException
     */
    public function anchor(string $cellColumn, int $cellRow)
    {
        if (Utility::validSheetCell($cellColumn, $cellRow)){
            $this->cellAnchor->height = $cellColumn;
            $this->cellAnchor->row   = $cellRow;
            $this->resolveAddresses();
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
