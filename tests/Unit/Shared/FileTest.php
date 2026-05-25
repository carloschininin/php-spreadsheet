<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Unit\Shared;

use CarlosChininin\Spreadsheet\Shared\File;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use PHPUnit\Framework\TestCase;

final class FileTest extends TestCase
{
    public function testItAppendsMissingSpreadsheetExtension(): void
    {
        self::assertSame('report.xlsx', File::updateFileName('report', SpreadsheetType::XLSX));
    }

    public function testItKeepsExistingSpreadsheetExtensionCaseInsensitively(): void
    {
        self::assertSame('report.XLSX', File::updateFileName('report.XLSX', SpreadsheetType::XLSX));
    }
}
