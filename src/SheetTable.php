<?php declare(strict_types=1);
namespace Crombi\PhpSpreadsheetHelper;

use \Exception;
/**
 * Class SheetTable
 * Collection of columns and tables with optional header and footer.
 *
 * A SheetTable is an ordered collection of columns, or tables, which are
 * groupings there of. Each one is rendered to the right of the previous one.
 * Optionally, a sheet table may also have a title and a footer. The title is
 * rendered above all columns, and the footer below.
 *
 * A table in turn consists of a title, the data table, and a summary footer
 * (e.g. totals). Of these only the data table is required.
 *
 * Style may be applied to particular cells, or ranges of. By default columns
 * are filled in, with default left-alignment for any non-numeric data, and
 * right-alignment for numeric data.
 *
 * A sheet tables width and height is variable. The width is the sum of the widths
 * of each column and sub-table, but the height is the height of the largest
 * column or sub-table.
 *
 * @todo For now columns must be constructed manually. Adding in values, be they simple sets,
 *       such as (1, 2, 3) or complex sets (1, (2, 3)) is currently pending. Each entry should
 *       be interpreted as a column, with subarray representing a sub-table.
 *
 * @todo Add being able to tag a column or table with an ID so the user can retrieve it
 *       later to alter properties.
 *
 * @todo Decouple classes from PHPSpreadsheet and its styling. Styling should probably
 *       be refactored to a Trait so as to gather as much externally dependant code
 *       into one place.
 *
 * @todo Support LEFT_RIGHT columns (building tables from 'rows').
 *
 * @todo Allow attaching data type metadata on different table entities. By attaching
 *       data type to collection of entities (table or columns), all entity elements
 *       in that collection are guaranteed to be of the given datatype.
 *
 * @todo Currently interrelation between elements is not allowed. Adding this in
 *       should be considered, particularly in terms of offsetting anchor points
 *       based on other elements (spatial interrelations).
 *
 * @link https://phpspreadsheet.readthedocs.io/en/develop/topics/recipes/#styles
 */
class SheetTable extends AnchorableEntity
{

    /**
     * @var array $tableElements
     */
    private $tableElements;
    private $header;
    private $footer;
    private $styleArray;

    /**
     * @var bool Specifies how a table will add new values. If the table is rectangular,
     *            new elements are always appended to the table. If not, new values are appended
     *            to the end of columns. This is similar in appearance to jagged arrays.
     */
    private $isRectangularTable;

    public function __construct()
    {
        parent::__construct();
        $this->setStyleArray([]);
        $this->tableElements = array();
    }

    /**
     * Adds an ordered set of columns to the tables column ordered collection.
     *
     * @param AnchorableEntity[] $tableEntities
     * @return SheetTable
     */
    public function addElements(AnchorableEntity ...$tableEntities) : self
    {
        foreach($tableEntities as $entity){
            if($tableEntities !== NULL) {
                $this->tableElements[] = $entity;
            }
        }

        return $this;
    }

    /**
     * @inheritdoc
     */
    protected function resolveAddresses() : self
    {
        $tableColumn = $this->getAnchor()->column;
        $tableRow = $this->getAnchor()->row;

        //header cell comes first
        if($this->header !== NULL) {
            $this->header->anchor($tableColumn, $tableRow);
            $tableRow += $this->getHeader()->getSheetCellHeight();
        }

        foreach ($this->tableElements as $element) {
            $element->anchor($tableColumn, $tableRow);
            for ($width = $element->getSheetCellWidth(); $width >= 1; $tableColumn++, $width--);
        }

        $lowestRow = $this->getLowerRightCell()->row;
        //footer cell comes last
        if($this->footer !== NULL) {
            $this->footer->anchor($this->getAnchor()->column, ++$lowestRow);
        }

        return $this;
    }

    /**
     * @param object|NULL $headerValue Sets header to given value. If header value
     *                                 is NULL, header is removed.
     *
     * @return SheetTable
     */
    public function setHeader($headerValue) : self
    {
        $headerCell = $this->header;

        try {
            if($headerValue !== NULL) {
                $headerCell = new SheetTableCell($headerValue);
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
     * @return SheetTable
     */
    public function setFooter($footerValue) : self
    {
        $footerCell = $this->footer;

        try {
            if($footerValue !== NULL) {
                $footerCell = new SheetTableCell($footerValue);
            }
        } catch (\Exception $e) {
            //Width exception should never occur, as if it does, it will occur
            //when setting the columns width
        }
        $this->footer = $footerCell;

        return $this;
    }

    /**
     * @todo
     * @param array[] $value
     *
     * @return SheetTable
     */
    public function addValues(array ...$valueArrays) : self
    {
        foreach($valueArrays as $array){
            $index = 0;
            for ($index = 0; $index < 3; $index++){
                //if ($index < count($this->getElements())) {
                if (isset($array[$index])) {
                    $this->getElements()[$index]->addValues([$array[$index]]);
                    //$index++;
                }else{
                    $this->getElements()[$index]->addValues(['']);
                }

            }
        }

        return $this;
    }

    /**
     * Returns the array of elements associated with this table.
     *
     * ___WARNING__: DO NOT RELY ON THIS CLASS EXPOSING ITS UNDERLYING ELEMENTS.
     *
     * @todo This is a patch added on the fly. Exposing TableSheetColumns is leaky
     *       design and a bad idea, as it opens up a whole host of edge cases
     *       where cell anchors and dimensions are edited _after_ being added
     *       to the container.
     *
     * @return array|null
     */
    public function getElements() : array
    {
        return $this->tableElements;
    }

    /**
     * @return bool
     */
    public function isRectangularTable() : bool
    {
        return $this->isRectangularTable;
    }

    /**
     * @param bool $isRectangular
     * @return SheetTable
     */
    public function rectangularTable(bool $isRectangular) : self
    {
        $this->isRectangularTable = $isRectangular;

        return $this;
    }

    /**
     * @return SheetTableCell|null
     */
    public function getHeader() : ?SheetTableCell
    {
        if($this->header !== NULL)
            return clone $this->header;

        return NULL;
    }

    /**
     * @return SheetTableCell|null
     */
    public function getFooter() : ?SheetTableCell
    {
        if($this->footer !== NULL)
            return clone $this->footer;

        return NULL;
    }
    /**
     * Sets the tables style. A tables style applies to all elements it contains.
     * It is important however to note, however, that an elements style takes
     * precedence over the tables style.
     *
     * @param array $styleArray
     *
     * @return SheetTable
     */
    public function setStyleArray (array $styleArray) : self
    {
        $this->styleArray = $styleArray;

        return $this;
    }

    /**
     * @return array
     */
    public function getStyleArray () : array
    {
        return $this->styleArray;
    }

    /**
     * @return array
     */
    public function getHeaderStyleArray() : array
    {
        if($this->getHeader() === NULL) {
            return NULL;
        }
        return $this->getHeader()->getStyleArray();
    }

    /**
     * @return array
     */
    public function getFooterStyleArray() : array
    {
        if($this->getFooter() === NULL){
            return NULL;
        }
        return $this->getFooter()->getStyleArray();
    }

    /**
     * @param array $styleArray
     * @return SheetTable
     */
    public function setHeaderStyleArray(array $styleArray) : self
    {
        if ($this->header !== NULL) {
        $this->header->setStyleArray($styleArray);
        }

        return $this;
    }

    /**
     * @param array $styleArray
     * @return SheetTable
     */
    public function setFooterStyleArray(array $styleArray) : self
    {
        if($this->getFooter() !== NULL){
            $this->footer->setStyleArray($styleArray);
        }
        return $this;
    }

    /**
     * Setting height of a table is not allowed.
     *
     * @param int $height
     *
     * @return SheetTable
     * @throws Exception
     */
    public function setSheetCellHeight(int $height) : self
    {
        throw new Exception();
    }

    /**
     * Setting width of a table is not allowed.
     *
     * @param int $width
     *
     * @return SheetTable
     * @throws Exception
     */
    public function setSheetCellWidth(int $width) : self
    {
        throw new Exception();
    }

    //TODO: Calculations need to be written

    /**
     * The height of a table is the greatest height among all of its elements.
     * @return int
     */
    public function getSheetCellHeight() : int
    {
        $heights = [0];

        if($this->tableElements !== []){
            foreach($this->tableElements as $element){
                $heights[] = $element->getSheetCellHeight();
            }
        }

        if($this->getHeader() !== NULL) {
            $heights[] = $this->getHeader()->getSheetCellHeight();
        }

        if($this->getFooter() !== NULL) {
            $heights[] = $this->getFooter()->getSheetCellHeight();
        }

        return max($heights);
    }

    /**
     * The width of a table is the sum of the widths of its elements.
     * @return int
     */
    public function getSheetCellWidth(): int
    {
        $widths = [0];

        if($this->tableElements !== []){
            foreach($this->tableElements as $element){
                $widths[] = $element->getSheetCellWidth();
            }
        }

        return array_sum($widths);
    }
}
