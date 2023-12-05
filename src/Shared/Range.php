<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Shared;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;

class Range
{
    public readonly string $value;
    public readonly int|string $firstColumn;
    public readonly int $firstRow;
    public readonly int|string $lastColumn;
    public readonly int $lastRow;

    private function __construct(string $range, $numberColumn = false)
    {
        if (!str_contains($range, ':')) {
            $range .= ":{$range}";
        }

        if (1 !== preg_match('/^([A-Z]+)(\\d+):([A-Z]+)(\\d+)$/', $range, $matches)) {
            throw new \RuntimeException('Not valid range of cells.');
        }

        $this->value = $range;
        $this->firstRow = (int) $matches[2];
        $this->lastRow = (int) $matches[4];
        try {
            $this->firstColumn = !$numberColumn ? $matches[1] : Coordinate::columnIndexFromString($matches[1]);
            $this->lastColumn = !$numberColumn ? $matches[3] : Coordinate::columnIndexFromString($matches[3]);
        } catch (Exception) {
            throw new \RuntimeException('Error convert columns to numbers');
        }
    }

    public function start(): string
    {
        return $this->firstColumn.$this->firstRow;
    }

    public function end(): string
    {
        return $this->lastColumn.$this->lastRow;
    }

    public static function create(string $range, $numberColumn = false): static
    {
        return new self($range, $numberColumn);
    }
}
