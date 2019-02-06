<?php declare(strict_type=1);
namespace Crombi\PhpSpreadsheetHelper;

use Throwable;

class InvalidSheetCellAddressException extends \Exception
{
    public function __construct(string $address = '', Throwable $previous = NULL)
    {
        $message = $address;
        if ($address !== '')
            $message = "Invalid Sheet Cell Address. Address given: " . $address;

        parent::__construct($message, 0, $previous);
    }
}