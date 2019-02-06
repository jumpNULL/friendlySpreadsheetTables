<?php declare(strict_type = 1);
namespace Crombi\PhpSpreadsheetHelper;

use Throwable;

class InvalidTableCellDimensionException extends \Exception
{
    public function __construct(string $dimension, int $size, Throwable $previous = NULL)
    {
        parent::__construct("Invalid table cell dimension given. " .
            $dimension . ": " . $size, 1, $previous);
    }
}