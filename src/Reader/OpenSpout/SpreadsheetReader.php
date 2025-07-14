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
use OpenSpout\Reader\Exception\ReaderNotOpenedException;
use OpenSpout\Reader\SheetInterface;

class SpreadsheetReader implements ReaderInterface
{
    use ReaderTrait;

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

            $this->reader->close();

            return $cells;
        } catch (IOException|ReaderNotOpenedException $e) {
            throw new ReaderException($e->getMessage());
        }
    }

    public function iterator(callable $callback): static
    {
        try {
            $this->reader->open($this->fileName);
            $sheetIterator = $this->reader->getSheetIterator();
            foreach ($sheetIterator as $sheet) {
                foreach ($sheet->getRowIterator() as $index => $row) {
                    $cells = array_map(fn (Cell $cell) => $cell->getValue(), $row->getCells());
                    \call_user_func($callback, $cells, $index);
                }

                break; // only first sheet
            }
            $this->reader->close();

            return $this;
        } catch (IOException|ReaderNotOpenedException $e) {
            throw new ReaderException($e->getMessage());
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
            SpreadsheetType::ODS => OdsReader::create($this->options),
            SpreadsheetType::CSV => CsvReader::create($this->options),
            default => XlsxReader::create($this->options),
        };
    }

    protected function getDataSheet(SheetInterface $sheet): array
    {
        $cells = [];
        foreach ($sheet->getRowIterator() as $row) {
            $cells[] = array_map(fn (Cell $cell) => $cell->getValue(), $row->getCells());
        }

        return $cells;
    }
}
