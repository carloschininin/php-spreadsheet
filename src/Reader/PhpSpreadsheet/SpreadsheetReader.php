<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader\PhpSpreadsheet;

use CarlosChininin\Spreadsheet\Reader\ReaderInterface;
use CarlosChininin\Spreadsheet\Reader\ReaderOptions;
use CarlosChininin\Spreadsheet\Reader\ReaderTrait;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use PhpOffice\PhpSpreadsheet\Reader\BaseReader;
use PhpOffice\PhpSpreadsheet\Reader\Csv as CsvReader;
use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;
use PhpOffice\PhpSpreadsheet\Reader\Ods as OdsReader;
use PhpOffice\PhpSpreadsheet\Reader\Xls as XlsReader;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class SpreadsheetReader implements ReaderInterface
{
    use ReaderTrait;

    protected BaseReader $reader;

    private function __construct(string $fileName, ReaderOptions $options)
    {
        $this->startReader($fileName, $options);
    }

    public static function create(string $fileName, ?ReaderOptions $options = null): static
    {
        return new self($fileName, $options ?? new ReaderOptions());
    }

    public function load(string $fileName, ?ReaderOptions $options = null): static
    {
        $this->startReader($fileName, $options);

        return $this;
    }

    public function data(bool $multipleSheet = false): array
    {
        $spreadsheet = null;
        try {
            $spreadsheet = $this->reader->load($this->fileName);
            if (!$multipleSheet) {
                return $spreadsheet->getActiveSheet()->toArray(formatData: false, ignoreHidden: true);
            }

            $cells = [];
            foreach ($spreadsheet->getAllSheets() as $sheet) {
                $cells[$sheet->getTitle()] = $sheet->toArray(formatData: false);
            }

            return $cells;
        } finally {
            $this->disconnectSpreadsheet($spreadsheet);
        }
    }

    public function iterator(callable $callback): static
    {
        $worksheetInfo = $this->reader->listWorksheetInfo($this->fileName);
        if ([] === $worksheetInfo || $worksheetInfo[0]['totalRows'] < 1) {
            return $this;
        }

        $chunkSize = max(1, $this->options?->readChunkSize ?? 1000);
        $firstSheet = $worksheetInfo[0];
        $totalRows = $firstSheet['totalRows'];
        $lastColumn = $firstSheet['lastColumnLetter'];
        $chunkFilter = new ChunkReadFilter();

        $readFilter = $this->reader->getReadFilter();
        $loadSheetsOnly = $this->reader->getLoadSheetsOnly();
        $readEmptyCells = $this->reader->getReadEmptyCells();

        try {
            $this->reader
                ->setReadFilter($chunkFilter)
                ->setLoadSheetsOnly($firstSheet['worksheetName'])
                ->setReadEmptyCells(false);

            for ($startRow = 1; $startRow <= $totalRows; $startRow += $chunkSize) {
                $endRow = min($startRow + $chunkSize - 1, $totalRows);
                $chunkFilter->setRows($startRow, $chunkSize);

                $spreadsheet = null;
                try {
                    $spreadsheet = $this->reader->load($this->fileName);
                    $sheet = $spreadsheet->getActiveSheet();
                    foreach ($sheet->rangeToArrayYieldRows("A{$startRow}:{$lastColumn}{$endRow}", null, false, false) as $offset => $cells) {
                        \call_user_func($callback, $cells, $startRow + $offset);
                    }
                } finally {
                    $this->disconnectSpreadsheet($spreadsheet);
                }
            }

            return $this;
        } finally {
            $this->restoreReaderState($readFilter, $loadSheetsOnly, $readEmptyCells);
        }
    }

    protected function startReader(string $fileName, ?ReaderOptions $options): void
    {
        $this->type = $this->getType($fileName);
        $this->fileName = $fileName;
        $this->options = $options ?? $this->options;
        $this->reader = match ($this->type) {
            SpreadsheetType::ODS => new OdsReader(),
            SpreadsheetType::CSV => $this->createCsvReader($this->options),
            SpreadsheetType::XLS => new XlsReader(),
            default => new XlsxReader(),
        };

        $this->reader->setReadDataOnly(true);
    }

    private function createCsvReader(?ReaderOptions $options): CsvReader
    {
        $reader = new CsvReader();
        if ($options?->fieldDelimiter) {
            $reader->setDelimiter($options->fieldDelimiter);
        }

        return $reader;
    }

    private function disconnectSpreadsheet(?Spreadsheet $spreadsheet): void
    {
        if (null !== $spreadsheet) {
            $spreadsheet->disconnectWorksheets();
        }
    }

    /**
     * @param string[]|null $loadSheetsOnly
     */
    private function restoreReaderState(IReadFilter $readFilter, ?array $loadSheetsOnly, bool $readEmptyCells): void
    {
        $this->reader
            ->setReadFilter($readFilter)
            ->setReadEmptyCells($readEmptyCells);

        if (null === $loadSheetsOnly) {
            $this->reader->setLoadAllSheets();

            return;
        }

        $this->reader->setLoadSheetsOnly($loadSheetsOnly);
    }
}
