<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Unit\Shared;

use CarlosChininin\Spreadsheet\Shared\Range;
use PHPUnit\Framework\TestCase;

final class RangeTest extends TestCase
{
    public function testItParsesCellRangeWithColumnLabels(): void
    {
        $range = Range::create('B2:C4');

        self::assertSame('B2:C4', $range->value);
        self::assertSame('B', $range->firstColumn);
        self::assertSame(2, $range->firstRow);
        self::assertSame('C', $range->lastColumn);
        self::assertSame(4, $range->lastRow);
        self::assertSame('B2', $range->start());
        self::assertSame('C4', $range->end());
    }

    public function testItExpandsSingleCellAsRange(): void
    {
        $range = Range::create('A1');

        self::assertSame('A1:A1', $range->value);
        self::assertSame('A1', $range->start());
        self::assertSame('A1', $range->end());
    }

    public function testItCanExposeNumericColumns(): void
    {
        $range = Range::create('B2:C4', true);

        self::assertSame(2, $range->firstColumn);
        self::assertSame(3, $range->lastColumn);
    }

    public function testItRejectsInvalidRanges(): void
    {
        $this->expectException(\RuntimeException::class);
        $this->expectExceptionMessage('Not valid range of cells.');

        Range::create('bad-range');
    }
}
