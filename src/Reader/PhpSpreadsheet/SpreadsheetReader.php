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
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Reader\BaseReader;
use PhpOffice\PhpSpreadsheet\Reader\Csv as CsvReader;
use PhpOffice\PhpSpreadsheet\Reader\Ods as OdsReader;
use PhpOffice\PhpSpreadsheet\Reader\Xls as XlsReader;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;

class SpreadsheetReader implements ReaderInterface
{
    use ReaderTrait;

    protected BaseReader $reader;

    private function __construct(string $fileName, ReaderOptions $options)
    {
        $this->startReader($fileName, $options);
    }

    public static function create(string $fileName, ReaderOptions $options = null): static
    {
        return new self($fileName, $options ?? new ReaderOptions());
    }

    public function load(string $fileName, ReaderOptions $options = null): static
    {
        $this->startReader($fileName, $options);

        return $this;
    }

    public function data(bool $multipleSheet = false): array
    {
        $spreadsheet = $this->reader->load($this->fileName);
        if (!$multipleSheet) {
            return $spreadsheet->getActiveSheet()->toArray(formatData: false, ignoreHidden: true);
        }

        $cells = [];
        foreach ($spreadsheet->getAllSheets() as $sheet) {
            $cells[$sheet->getTitle()] = $sheet->toArray(formatData: false);
        }

        return $cells;
    }

    public function iterator(callable $callback): static
    {
        $spreadsheet = $this->reader->load($this->fileName);

        foreach ($spreadsheet->getWorksheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $index => $row) {
                $cells = array_map(fn (Cell $cell) => $cell->getValue(), iterator_to_array($row->getCellIterator(), false));
                \call_user_func($callback, $cells, $index);
            }

            break; // only first sheet
        }

        return $this;
    }

    protected function startReader(string $fileName, ?ReaderOptions $options): void
    {
        $this->type = $this->getType($fileName);
        $this->fileName = $fileName;
        $this->options = $options ?? $this->options;
        $this->reader = match ($this->type) {
            SpreadsheetType::ODS => new OdsReader(), // Failed
            SpreadsheetType::CSV => new CsvReader(),
            SpreadsheetType::XLS => new XlsReader(),
            default => new XlsxReader(),
        };

        $this->reader->setReadDataOnly(true);
    }
}
