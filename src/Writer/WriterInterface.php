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

    public function setCellValue(string|int $col, int $row, mixed $value, DataFormat|string|null $format = null, ?DataType $type = null): static;

    // Only PHPSpreadsheet
    public function columnAutoSize(string|int|null $start = null, string|int|null $end = null): static;

    public function mergeCells(string|array $start, string|array $end, mixed $value = null, ?array $style = null): static;

    public function styleCells(string|array $start, string|array $end, array $style): static;

    public function fromArray(string|int $col, int $row, mixed $data, mixed $style = null): static;

    public function formatCells(DataFormat|string $format, string|array $start, string|array|null $end = null): static;

    public function addSheet(string $title, bool $isActive = true): static;

    public function activeSheet(string $title): static;

    public function removeInitialSheet(): static;

    public function renameSheet(int|string $sheetIndexOrTitle, string $newTitle): bool;
}
