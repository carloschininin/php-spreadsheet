<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader;

final class ReaderException extends \RuntimeException
{
    public function __construct($message = 'Spreadsheet reader exception')
    {
        parent::__construct(sprintf('Error: %s', $message));
    }
}
