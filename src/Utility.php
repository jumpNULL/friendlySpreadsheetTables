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
     * @param string $row
     *
     * @return bool
     */
    public static function validSheetCell(string $column, $row) : bool
    {
        $columnRegex = '/^[A-Za-z]+$/';
        $rowRegex = '/^[1-9][0-9]*$/';

        return preg_match($rowRegex, $row) &&
            preg_match($columnRegex, $column);
    }
}
