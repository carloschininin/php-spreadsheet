<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Unit\Writer\OpenSpout;

use CarlosChininin\Spreadsheet\Shared\DataType;
use CarlosChininin\Spreadsheet\Writer\OpenSpout\Helper;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;
use OpenSpout\Common\Entity\Cell\DateTimeCell;
use OpenSpout\Common\Entity\Cell\EmptyCell;
use OpenSpout\Common\Entity\Cell\NumericCell;
use OpenSpout\Common\Entity\Cell\StringCell;
use PHPUnit\Framework\TestCase;

final class HelperTest extends TestCase
{
    public function testItConvertsConfiguredDataTypesBeforeBuildingCells(): void
    {
        $options = new WriterOptions(dataTypes: ['amount' => DataType::FLOAT]);

        $cell = Helper::valueToCell('12.50', $options, 'amount');

        self::assertInstanceOf(NumericCell::class, $cell);
        self::assertSame(12.5, $cell->getValue());
    }

    public function testItBuildsDedicatedCellsForKnownValues(): void
    {
        self::assertInstanceOf(DateTimeCell::class, Helper::valueToCell(new \DateTimeImmutable('2026-01-01'), new WriterOptions(), 0));
        self::assertInstanceOf(NumericCell::class, Helper::valueToCell(10.25, new WriterOptions(), 0));

        $boolCell = Helper::valueToCell(true, new WriterOptions(), 0);
        self::assertInstanceOf(StringCell::class, $boolCell);
        self::assertSame('SI', $boolCell->getValue());
    }

    public function testItConvertsStringableObjectsAndIgnoresPlainObjects(): void
    {
        $stringable = new class {
            public function __toString(): string
            {
                return 'object-value';
            }
        };

        $stringCell = Helper::valueToCell($stringable, new WriterOptions(), 0);
        $emptyCell = Helper::valueToCell(new \stdClass(), new WriterOptions(), 0);

        self::assertInstanceOf(StringCell::class, $stringCell);
        self::assertSame('object-value', $stringCell->getValue());
        self::assertInstanceOf(EmptyCell::class, $emptyCell);
        self::assertNull($emptyCell->getValue());
    }
}
