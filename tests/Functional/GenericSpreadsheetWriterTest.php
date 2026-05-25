<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Tests\Functional;

use CarlosChininin\Spreadsheet\Reader\OpenSpout\SpreadsheetReader as OpenSpoutReader;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use CarlosChininin\Spreadsheet\Writer\SpreadsheetWriter;
use CarlosChininin\Spreadsheet\Writer\WriterBackend;
use CarlosChininin\Spreadsheet\Writer\WriterException;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;

final class GenericSpreadsheetWriterTest extends FunctionalTestCase
{
    public function testItUsesFastestAvailableBackendForSimpleXlsxExports(): void
    {
        $writer = new SpreadsheetWriter(
            data: [[1, 'Ana']],
            headers: ['ID', 'Nombre'],
            options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir),
        );

        $response = $writer->execute()->download('generic-simple');
        $file = $this->fileFromResponse($response, 'xlsx');
        $rows = OpenSpoutReader::create($file)->data();

        $expectedBackend = \extension_loaded('xlswriter') && class_exists(\Vtiful\Kernel\Excel::class)
            ? WriterBackend::ExcelWriter
            : WriterBackend::OpenSpout;

        self::assertSame($expectedBackend, $writer->backend());
        self::assertSame(['ID', 'Nombre'], $rows[0]);
        self::assertSame([1, 'Ana'], $rows[1]);
    }

    public function testItFallsBackToPhpSpreadsheetForXlsExports(): void
    {
        $writer = new SpreadsheetWriter(
            data: [[1, 'Ana']],
            headers: ['ID', 'Nombre'],
            options: new WriterOptions(type: SpreadsheetType::XLS, path: $this->tempDir),
        );

        $writer->execute()->download('generic-xls');

        self::assertSame(WriterBackend::PhpSpreadsheet, $writer->backend());
    }

    public function testAdvancedOperationsBeforeExecutionForcePhpSpreadsheetBackend(): void
    {
        $writer = new SpreadsheetWriter(
            data: [[1, 'Ana']],
            headers: ['ID', 'Nombre'],
            options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir),
        );

        $writer
            ->columnAutoSize('A', 'B')
            ->execute()
            ->download('generic-advanced');

        self::assertSame(WriterBackend::PhpSpreadsheet, $writer->backend());
    }

    public function testAdvancedOperationsAfterFastBackendSelectionFailExplicitly(): void
    {
        $writer = new SpreadsheetWriter(
            data: [[1, 'Ana']],
            headers: ['ID', 'Nombre'],
            options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir),
        );
        $writer->execute();

        $this->expectException(WriterException::class);
        $this->expectExceptionMessage('Operation "columnAutoSize" is not supported by selected');

        $writer->columnAutoSize('A', 'B');
    }

    public function testFromArrayUsesOpenSpoutWhenAvailableInsteadOfExcelWriter(): void
    {
        $writer = new SpreadsheetWriter(options: new WriterOptions(type: SpreadsheetType::XLSX, path: $this->tempDir));
        $writer->fromArray('A', 1, [['Clave', 'Valor'], ['total', 2]]);

        $response = $writer->download('generic-from-array');
        $file = $this->fileFromResponse($response, 'xlsx');
        $rows = OpenSpoutReader::create($file)->data();

        self::assertSame(WriterBackend::OpenSpout, $writer->backend());
        self::assertSame(['Clave', 'Valor'], $rows[0]);
        self::assertSame(['total', 2], $rows[1]);
    }
}
