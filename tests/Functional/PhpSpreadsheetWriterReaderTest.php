<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Functional;

use CarlosChininin\Spreadsheet\Reader\PhpSpreadsheet\SpreadsheetReader;
use CarlosChininin\Spreadsheet\Reader\ReaderOptions;
use CarlosChininin\Spreadsheet\Shared\DataType;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use CarlosChininin\Spreadsheet\Writer\PhpSpreadsheet\SpreadsheetWriter;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as PhpSpreadsheetDocument;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

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

    public function testItWritesConfiguredDataTypesWithPhpSpreadsheetBackend(): void
    {
        $writer = new SpreadsheetWriter(
            data: [
                ['001', '12.50', '7'],
            ],
            headers: ['Codigo', 'Decimal', 'Entero'],
            options: new WriterOptions(
                type: SpreadsheetType::XLSX,
                dataTypes: [0 => DataType::STRING, 1 => DataType::FLOAT, 2 => DataType::INT],
                path: $this->tempDir,
            ),
        );

        $response = $writer->execute(false)->download('typed-values');
        $file = $this->fileFromResponse($response, 'xlsx');
        $rows = SpreadsheetReader::create($file)->data();

        self::assertSame('001', $rows[1][0]);
        self::assertSame(12.5, $rows[1][1]);
        self::assertSame(7, $rows[1][2]);
    }

    public function testItStylesFromArrayWhenColumnIsString(): void
    {
        $writer = new SpreadsheetWriter(options: new WriterOptions(path: $this->tempDir));
        $writer->fromArray('A', 1, [['Clave', 'Valor']], ['font' => ['bold' => true]]);

        $response = $writer->download('styled-array');
        $file = $this->fileFromResponse($response, 'xlsx');
        $spreadsheet = IOFactory::load($file);

        try {
            self::assertTrue($spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->getBold());
            self::assertTrue($spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->getBold());
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    public function testItAutosizesMultiLetterColumns(): void
    {
        $writer = new SpreadsheetWriter(headers: array_fill(0, 30, 'h'));
        $writer->columnAutoSize('AA', 'AB');

        $property = new \ReflectionProperty($writer, 'writer');
        $spreadsheet = $property->getValue($writer);
        $sheet = $spreadsheet->getActiveSheet();

        self::assertFalse($sheet->getColumnDimension('A')->getAutoSize());
        self::assertTrue($sheet->getColumnDimension('AA')->getAutoSize());
        self::assertTrue($sheet->getColumnDimension('AB')->getAutoSize());
    }

    public function testPhpSpreadsheetReaderUsesConfiguredCsvDelimiter(): void
    {
        $file = $this->tempDir.'/pipe.csv';
        file_put_contents($file, "clave|valor\ntotal|2\n");

        $rows = SpreadsheetReader::create($file, new ReaderOptions(fieldDelimiter: '|'))->data();

        self::assertSame(['clave', 'valor'], $rows[0]);
        self::assertSame(['total', 2], $rows[1]);
    }

    public function testPhpSpreadsheetReaderIteratorReadsRowsInChunks(): void
    {
        $spreadsheet = new PhpSpreadsheetDocument();
        $spreadsheet->getActiveSheet()->fromArray(
            [
                ['ID', 'Nombre', 'Extra'],
                [1, 'Ana', null],
                [2, null, 'X'],
                [3, 'Luis', 'Y'],
            ],
            startCell: 'A1',
            strictNullComparison: true,
        );

        $file = $this->tempDir.'/chunked-reader.xlsx';
        (new Xlsx($spreadsheet))->save($file);
        $spreadsheet->disconnectWorksheets();

        $rows = [];
        $indexes = [];
        $reader = SpreadsheetReader::create($file, new ReaderOptions(readChunkSize: 2));

        $reader->iterator(static function (array $cells, int $index) use (&$rows, &$indexes): void {
            $indexes[] = $index;
            $rows[] = $cells;
        });

        self::assertSame([1, 2, 3, 4], $indexes);
        self::assertSame(
            [
                ['ID', 'Nombre', 'Extra'],
                [1, 'Ana', null],
                [2, null, 'X'],
                [3, 'Luis', 'Y'],
            ],
            $rows,
        );

        self::assertSame(['ID', 'Nombre', 'Extra'], $reader->data()[0]);
    }
}
