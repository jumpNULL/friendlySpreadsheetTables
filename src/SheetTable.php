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
 */
class SheetTable {
    /**
     * @var array $column Map from column names to column configuration.
     */
    private $column = array();

    /**
     * @var bool $lockedColumns Indicates whether columns are locked, preventing
     *   new columns from being added.
     */
    private $lockedColumns = false;

    public function __construct()
    {
    }

    /**
     * Adds an ordered set of columns to the tables column ordered collection.
     *
     * @param array $columns Set of columns to insert into table.
     * @return SheetTable
     */
    public function addColumns(array $columns) {
        return $this;
    }

    /**
     * Adds a row and optional cell-associated style. Cell styles take priority
     * over the general column style.
     *
     * @param array $rows Set of values to insert into table.
     * @return SheetTable
     */
    public function addRows(array $rows){
        return $this;
    }
}