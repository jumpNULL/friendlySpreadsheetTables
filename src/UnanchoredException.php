<?php

namespace Crombi\PhpSpreadsheetHelper;

use Throwable;

class UnanchoredException extends \Exception
{
    public function __construct(string $message = "", int $code = 0, Throwable $previous = NULL)
    {
        parent::__construct($message, $code, $previous);
    }
}
