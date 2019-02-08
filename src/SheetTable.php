<?php declare(strict_types=1);
namespace Crombi\PhpSpreadsheetHelper;

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
 *
 * @link https://phpspreadsheet.readthedocs.io/en/develop/topics/recipes/#styles
 */
class SheetTable extends AnchorableEntity
{

    /**
     * @var array $column Map from column names to column configuration.
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
    }

    /**
     * Adds an ordered set of columns to the tables column ordered collection.
     *
     * @param AnchorableEntity[] $tableEntities
     * @return SheetTable
     */
    public function addColumns(AnchorableEntity ...$tableEntities) : SheetTable
    {
        if(!is_null($tableEntities))
            array_merge($this->tableElements, $tableEntities);

        return $this;
    }

    /**
     * @inheritdoc
     */
    public function resolveAddresses()
    {
        $cellColumn = $this->cellAnchor->column;
        $cellRow = $this->cellAnchor->row;

        //header cell comes first
        if(!is_null($this->header)) {
            $this->header->anchor($cellColumn, $cellRow);
            $cellRow += $this->header->getSheetCellWidth();
        }

        foreach ($this->tableElements as $cell) {
            $cell->anchor($cellColumn, $cellRow);
            $cellRow += $cell->getSheetCellWidth();
        }

        //footer cell comes last
        if(!is_null($this->footer))
            $this->footer->anchor($cellColumn, $cellRow);
    }

    /**
     *
     * @param array[] $value
     *
     * @return SheetTable
     */
    public function addValues(array ...$value) : SheetTable
    {
        return $this;
    }

    public function isRectangularTable() : bool
    {
        return $this->isRectangularTable;
    }

    public function rectangularTable(bool $isRectangular) : SheetTable
    {
        $this->isRectangularTable = $isRectangular;

        return $this;
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
    public function setStyleArray (array $styleArray) : SheetTable
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
     * Applying a style to a column applies its style to every cell it contains.
     * It is important to note, however, that a cells style takes priority over
     * a columns style.
     *
     * @return SheetTable
     */
    public function applyStyle() : SheetTable
    {
        return $this;
    }
}
