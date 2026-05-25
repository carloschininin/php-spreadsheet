<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Functional;

use CarlosChininin\Spreadsheet\Reader\PhpSpreadsheet\SpreadsheetReader;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use CarlosChininin\Spreadsheet\Writer\PhpSpreadsheet\SpreadsheetWriter;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;

final class PhpSpreadsheetWriterReaderTest extends FunctionalTestCase
{
    public function testItWritesAndReadsXlsxDataWithPhpSpreadsheetBackend(): void
    {
        $writer = new SpreadsheetWriter(
            data: [
                [1, 'Ana', true, 12.5],
                [2, 'Luis', false, 8.75],
            ],
            headers: ['ID', 'Nombre', 'Activo', 'Total'],
            options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir, nameSheet: 'Data'),
        );

        $response = $writer->execute(false)->download('people');
        $file = $this->fileFromResponse($response, 'xlsx');

        $rows = SpreadsheetReader::create($file)->data();

        self::assertSame(['ID', 'Nombre', 'Activo', 'Total'], $rows[0]);
        self::assertSame([1, 'Ana', 'SI', 12.5], $rows[1]);
        self::assertSame([2, 'Luis', 'NO', 8.75], $rows[2]);
    }

    public function testItWritesAndReadsMultipleSheets(): void
    {
        $writer = new SpreadsheetWriter(options: new WriterOptions(path: $this->tempDir, nameSheet: 'Resumen'));
        $writer
            ->fromArray('A', 1, [['Clave', 'Valor'], ['total', 2]])
            ->addSheet('Detalle')
            ->fromArray('A', 1, [['ID', 'Nombre'], [1, 'Ana']]);

        $response = $writer->download('multi-sheet');
        $file = $this->fileFromResponse($response, 'xlsx');

        $sheets = SpreadsheetReader::create($file)->data(true);

        self::assertArrayHasKey('Resumen', $sheets);
        self::assertArrayHasKey('Detalle', $sheets);
        self::assertSame(['Clave', 'Valor'], $sheets['Resumen'][0]);
        self::assertSame(['ID', 'Nombre'], $sheets['Detalle'][0]);
    }
}
