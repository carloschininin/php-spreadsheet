<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Unit\Shared;

use CarlosChininin\Spreadsheet\Shared\Column;
use PHPUnit\Framework\TestCase;

final class ColumnTest extends TestCase
{
    public function testItConvertsAsciiColumnNumberToExcelLabel(): void
    {
        self::assertSame('A', Column::numberToLabel(\ord('A')));
        self::assertSame('Z', Column::numberToLabel(\ord('Z')));
        self::assertSame('AA', Column::numberToLabel(\ord('Z') + 1));
        self::assertSame('AB', Column::numberToLabel(\ord('Z') + 2));
    }

    public function testItConvertsExcelLabelToZeroBasedNumber(): void
    {
        self::assertSame(0, Column::labelToNumber('A'));
        self::assertSame(25, Column::labelToNumber('Z'));
        self::assertSame(26, Column::labelToNumber('AA'));
        self::assertSame(27, Column::labelToNumber('ab'));
    }
}
