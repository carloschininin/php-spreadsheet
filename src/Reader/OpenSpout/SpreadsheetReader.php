<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader\OpenSpout;

use CarlosChininin\Spreadsheet\Reader\ReaderException;
use CarlosChininin\Spreadsheet\Reader\ReaderInterface;
use CarlosChininin\Spreadsheet\Reader\ReaderOptions;
use CarlosChininin\Spreadsheet\Reader\ReaderTrait;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Reader\AbstractReader;
use OpenSpout\Reader\CSV\Options as CsvOptions;
use OpenSpout\Reader\CSV\Reader as OpenSpoutCsvReader;
use OpenSpout\Reader\Exception\ReaderNotOpenedException;
use OpenSpout\Reader\ODS\Options as OdsOptions;
use OpenSpout\Reader\ODS\Reader as OdsReader;
use OpenSpout\Reader\SheetInterface;
use OpenSpout\Reader\XLSX\Options as XlsxOptions;
use OpenSpout\Reader\XLSX\Reader as XlsxReader;

class SpreadsheetReader implements ReaderInterface
{
    use ReaderTrait;

    private const string DEFAULT_CSV_DELIMITER = ';';

    protected AbstractReader $reader;

    private function __construct(string $fileName, ?ReaderOptions $options)
    {
        $this->startReader($fileName, $options);
    }

    public static function create(string $fileName, ?ReaderOptions $options = null): static
    {
        return new self($fileName, $options);
    }

    public function load(string $fileName, ?ReaderOptions $options = null): static
    {
        $this->startReader($fileName, $options);

        return $this;
    }

    public function data(bool $multipleSheet = false): array
    {
        try {
            $this->reader->open($this->fileName);
            $sheetIterator = $this->reader->getSheetIterator();
            $cells = [];
            foreach ($sheetIterator as $sheet) {
                if (!$multipleSheet) {
                    $cells = $this->getDataSheet($sheet);
                    break;
                }

                $cells[$sheet->getName()] = $this->getDataSheet($sheet);
            }

            return $cells;
        } catch (IOException|ReaderNotOpenedException $e) {
            throw new ReaderException($e->getMessage());
        } finally {
            $this->reader->close();
        }
    }

    public function iterator(callable $callback): static
    {
        try {
            $this->reader->open($this->fileName);
            $sheetIterator = $this->reader->getSheetIterator();
            foreach ($sheetIterator as $sheet) {
                foreach ($sheet->getRowIterator() as $index => $row) {
                    $cells = array_map(static fn (Cell $cell) => $cell->getValue(), $this->getRowCells($row));
                    \call_user_func($callback, $cells, $index);
                }

                break; // only first sheet
            }

            return $this;
        } catch (IOException|ReaderNotOpenedException $e) {
            throw new ReaderException($e->getMessage());
        } finally {
            $this->reader->close();
        }
    }

    protected function startReader(string $fileName, ?ReaderOptions $options): void
    {
        $this->type = $this->getType($fileName);
        if (SpreadsheetType::XLS === $this->type) {
            throw new ReaderException('Format XLS not support, use format XLSX');
        }

        $this->fileName = $fileName;
        $this->options = $options ?? $this->options;
        $this->reader = match ($this->type) {
            SpreadsheetType::ODS => new OdsReader(new OdsOptions()),
            SpreadsheetType::CSV => $this->createCsvReader($this->options),
            default => new XlsxReader(new XlsxOptions()),
        };
    }

    protected function getDataSheet(SheetInterface $sheet): array
    {
        $cells = [];
        foreach ($sheet->getRowIterator() as $row) {
            $cells[] = array_map(static fn (Cell $cell) => $cell->getValue(), $this->getRowCells($row));
        }

        return $cells;
    }

    private function createCsvReader(?ReaderOptions $options): OpenSpoutCsvReader
    {
        return new OpenSpoutCsvReader(new CsvOptions(
            SHOULD_PRESERVE_EMPTY_ROWS: true,
            FIELD_DELIMITER: $options?->fieldDelimiter ?? self::DEFAULT_CSV_DELIMITER,
        ));
    }

    private function getRowCells(object $row): array
    {
        if (method_exists($row, 'getCells')) {
            return $row->getCells();
        }

        return $row->cells;
    }
}
