<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer;

use CarlosChininin\Spreadsheet\Shared\DataFormat;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;

class WriterOptions
{
    public function __construct(
        public readonly SpreadsheetType $type = SpreadsheetType::XLSX,
        public readonly string $col = 'A',
        public readonly int $row = 1,
        public ?string $endCol = null,
        public ?int $numCols = null,
        public readonly bool $headerStyle = true,
        public readonly DataFormat $formatDate = DataFormat::DATE_DMYSLASH,
        public readonly DataFormat $formatDecimal = DataFormat::NUMBER_00,
        public readonly ?string $path = null,
    ) {
    }
}
