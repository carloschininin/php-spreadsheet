<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Functional;

use CarlosChininin\Spreadsheet\Reader\OpenSpout\SpreadsheetReader;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use CarlosChininin\Spreadsheet\Writer\OpenSpout\SpreadsheetWriter;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;

final class OpenSpoutWriterReaderTest extends FunctionalTestCase
{
    public function testItWritesAndReadsXlsxDataWithOpenSpoutBackend(): void
    {
        $writer = new SpreadsheetWriter(
            data: [
                [1, 'Ana', true, 12.5],
                [2, 'Luis', false, 8.75],
            ],
            headers: ['ID', 'Nombre', 'Activo', 'Total'],
            options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir),
        );

        $response = $writer->execute()->download('open-spout');
        $file = $this->fileFromResponse($response, 'xlsx');

        $rows = SpreadsheetReader::create($file)->data();

        self::assertSame(['ID', 'Nombre', 'Activo', 'Total'], $rows[0]);
        self::assertSame([1, 'Ana', 'SI', 12.5], $rows[1]);
        self::assertSame([2, 'Luis', 'NO', 8.75], $rows[2]);
    }
}
