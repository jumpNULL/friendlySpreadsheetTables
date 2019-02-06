<?php
namespace Crombi\PhpSpreadsheetHelper;
/**
 * Class Utility
 * Collection of function for validation and misc. utilized by other classes.
 *
 * @package Crombi\PhpSpreadsheetHelper
 */
class Utility
{
    /**
     * Validates if a given column and row are valid cell address.
     *
     * @todo This provides basic validation in terms of format, but does not validate
     *       that the address is actually within a containers address.
     *
     * @param string $column
     * @param        $row
     *
     * @return bool
     */
    public static function validSheetCell(string $column, $row) : bool
    {
        $rowRegex = '^[a-Z]+$';
        $columnRegex = '^[1-9][0-9]*$';

        if (preg_match($rowRegex, $row) === 1 &&
            preg_match($columnRegex, $column) === 1) {
            return true;
        } else
            return false;
    }
}