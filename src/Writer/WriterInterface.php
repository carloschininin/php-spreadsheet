<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer;

use CarlosChininin\Spreadsheet\Shared\DataFormat;
use CarlosChininin\Spreadsheet\Shared\DataType;
use Symfony\Component\HttpFoundation\Response;

interface WriterInterface
{
    public function execute(bool $direct = true): static;
    public function download(string $fileName): Response;
    public function col(): string;
    public function row(): int;
    public function numCols(): int;
    public function endCol(): string;
    public function range(): string;

    public function setCellValue(string|int $col, int $row, mixed $value, DataFormat $format = null, DataType $type = null): static;

    // Only PHPSpreadsheet
    public function columnAutoSize(string $start = null, string $end = null): static;
}
