<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Functional;

use CarlosChininin\Spreadsheet\Reader\OpenSpout\SpreadsheetReader;
use CarlosChininin\Spreadsheet\Reader\ReaderOptions;
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

    public function testItUsesConfiguredSheetNameWhenExecuting(): void
    {
        $writer = new SpreadsheetWriter(
            data: [[1, 'Ana']],
            headers: ['ID', 'Nombre'],
            options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir, nameSheet: 'Personas'),
        );

        $response = $writer->execute()->download('open-spout-named-sheet');
        $file = $this->fileFromResponse($response, 'xlsx');
        $sheets = SpreadsheetReader::create($file)->data(true);

        self::assertArrayHasKey('Personas', $sheets);
        self::assertSame(['ID', 'Nombre'], $sheets['Personas'][0]);
    }

    public function testItWritesAndReadsMultipleSheetsWithOpenSpoutBackend(): void
    {
        $writer = new SpreadsheetWriter(options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir, nameSheet: 'Resumen'));
        $writer
            ->fromArray('A', 1, [['Clave', 'Valor'], ['total', 2]])
            ->addSheet('Detalle')
            ->fromArray('A', 1, [['ID', 'Nombre'], [1, 'Ana']]);

        $response = $writer->download('open-spout-multi-sheet');
        $file = $this->fileFromResponse($response, 'xlsx');
        $sheets = SpreadsheetReader::create($file)->data(true);

        self::assertArrayHasKey('Resumen', $sheets);
        self::assertArrayHasKey('Detalle', $sheets);
        self::assertSame(['Clave', 'Valor'], $sheets['Resumen'][0]);
        self::assertSame(['ID', 'Nombre'], $sheets['Detalle'][0]);
    }

    public function testOpenSpoutReaderReadsCsvWithoutExplicitOptions(): void
    {
        $file = $this->tempDir.'/semicolon.csv';
        file_put_contents($file, "clave;valor\ntotal;2\n");

        $rows = SpreadsheetReader::create($file)->data();

        self::assertSame(['clave', 'valor'], $rows[0]);
        self::assertSame(['total', '2'], $rows[1]);
    }

    public function testOpenSpoutReaderUsesConfiguredCsvDelimiter(): void
    {
        $file = $this->tempDir.'/pipe.csv';
        file_put_contents($file, "clave|valor\ntotal|2\n");

        $rows = SpreadsheetReader::create($file, new ReaderOptions(fieldDelimiter: '|'))->data();

        self::assertSame(['clave', 'valor'], $rows[0]);
        self::assertSame(['total', '2'], $rows[1]);
    }

    public function testOpenSpoutReaderAcceptsOptionsForXlsxWithoutTypeError(): void
    {
        $writer = new SpreadsheetWriter(
            data: [[1, 'Ana']],
            headers: ['ID', 'Nombre'],
            options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir),
        );

        $response = $writer->execute()->download('open-spout-options');
        $file = $this->fileFromResponse($response, 'xlsx');
        $rows = SpreadsheetReader::create($file, new ReaderOptions())->data();

        self::assertSame(['ID', 'Nombre'], $rows[0]);
        self::assertSame([1, 'Ana'], $rows[1]);
    }
}
