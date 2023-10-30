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
    public readonly int $row;
    public readonly string $col;
    public ?string $endCol;

    public function __construct(
        public readonly SpreadsheetType $type = SpreadsheetType::XLSX,
        int $row = 1,
        string $col = 'A',
        string $endCol = null,
        public ?int $numCols = null,
        public readonly bool $headerStyle = true,
        public readonly DataFormat $formatDate = DataFormat::DATE_DMYSLASH,
        public readonly DataFormat $formatDecimal = DataFormat::NUMBER_00,
        public readonly ?string $path = null,
    ) {
        $this->row = $row > 0 ? $row : 1;
        $this->col = mb_strtoupper($col);
        $this->endCol = mb_strtoupper($endCol);
    }
}
