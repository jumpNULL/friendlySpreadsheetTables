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
    public function getLowerRightCellAddress() : string
    {
        if ($this->getSheetCellWidth() == 1 &&
            $this->getSheetCellHeight() == 1)
            return $this->getAnchorAddress();
        else {
             $lowerRightCell = $this->getLowerRightCell();
            //TODO: This will break if the column string is not a simple character.
            return $lowerRightCell->column . strval($lowerRightCell->row);
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
            $this->cellAnchor->column = $cellColumn;
            $this->cellAnchor->row   = $cellRow;
            $this->resolveAddresses();
        } else
            throw new InvalidSheetCellAddressException($cellColumn .
                strval($cellRow));

        return $this;
    }

    /**
     * @return string|NULL
     */
    public function getAnchorAddress() : string
    {
        if (!is_null($this->cellAnchor)) {
            return $this->cellAnchor->column . $this->cellAnchor->row;
        } else
            return NULL;
    }

    /**
     * @return object
     */
    public function getAnchor() : object
    {
        return clone $this->cellAnchor;
    }

    public function getLowerRightCell() : object
    {
        if(!$this->isAnchored())
            throw new UnanchoredException();

        if ($this->getSheetCellWidth() == 1 &&
            $this->getSheetCellHeight() == 1)
            return clone $this->cellAnchor;
        else {
            $width = $this->getSheetCellWidth();
            $height = $this->getSheetCellHeight();

            $cellColumn = $this->cellAnchor->column;
            for (; $width > 1; $cellColumn++, $width--) ;

            $cellRow = $this->cellAnchor->row + --$height;

            return (object) array(
                'column' => $cellColumn,
                'row' => $cellRow
            );
        }
    }
    /**
     * Returns whether a table element has been anchored to a sheet cell.
     *
     * @return bool
     */
    public function isAnchored() : bool
    {
        if (is_null($this->cellAnchor->column) || is_null($this->cellAnchor->row))
            return false;
        else
            return true;
    }
}
