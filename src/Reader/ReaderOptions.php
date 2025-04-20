<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader;

readonly class ReaderOptions
{
    public function __construct(
        public ?string $fieldDelimiter = null,
    ) {
    }
}
