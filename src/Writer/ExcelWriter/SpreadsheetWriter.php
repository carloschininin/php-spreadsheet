<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer\ExcelWriter;

use CarlosChininin\Spreadsheet\Shared\Column;
use CarlosChininin\Spreadsheet\Shared\DataFormat;
use CarlosChininin\Spreadsheet\Shared\DataType;
use CarlosChininin\Spreadsheet\Shared\File;
use CarlosChininin\Spreadsheet\Writer\WriterException;
use CarlosChininin\Spreadsheet\Writer\WriterInterface;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;
use CarlosChininin\Spreadsheet\Writer\WriterTrait;
use Symfony\Component\HttpFoundation\Response;
use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

/**
 * Spreadsheet writer fast
 */
class SpreadsheetWriter implements WriterInterface
{
    use WriterTrait;

    protected Excel $writer;

    public function __construct(
        protected readonly iterable $data = [],
        protected readonly iterable $headers = [],
        protected readonly WriterOptions $options = new WriterOptions(),
    ) {
        $this->createExcel();
    }

    public function execute(bool $direct = true): static
    {
        return match ($direct) {
            true => $this->executeDirect(),
            false => $this->executeCustom(),
        };
    }

    public function setCellValue(string|int $col, int $row, mixed $value, DataFormat $format = null, DataType $type = null): static
    {
        if (null === $value) {
            return $this;
        }

        if ($value instanceof \DateTimeInterface) {
            $value = $value->getTimestamp();
            $type = DataType::DATE;
            $format = $format ?? $this->options->formatDate;
        }

        if (\is_float($value)) {
            $type = DataType::NUMERIC;
            $format = $format ?? $this->options->formatDecimal;
        }

        if (\is_string($col)) {
            $col = Column::labelToNumber($col);
        }

        $this->writer = match ($type) {
            DataType::DATE => $this->writer->insertDate($row, $col, $value, $format?->value),
            DataType::FORMULA => $this->writer->insertFormula($row, $col, $value),
            default => $this->writer->insertText($row, $col, $value, $format?->value),
        };

        return $this;
    }

    public function mergeCells(string|array $start, string|array $end, mixed $value = null, array $style = null): static
    {
        // No implement in this library
        return $this;
    }

    public function styleCells(string|array $start, string|array $end, array $style): static
    {
        // No implement in this library
        return $this;
    }

    public function download(string $fileName, bool $useZip = false): Response
    {
        $fileName = File::updateFileName($fileName, $this->options->type);

        return $useZip
            ? File::downloadZip($fileName, $this->writer->output())
            : File::download($fileName, $this->writer->output());
    }

    public function columnAutoSize(string|int $start = null, string|int $end = null): static
    {
        // No implement in this library
        return $this;
    }

    public function fromArray(string|int $col, int $row, mixed $data, mixed $style = null): static
    {
        // No implement in this library
        return $this;
    }

    public function formatCells(DataFormat $format, string|array $start, string|array $end = null): static
    {
        // No implement in this library
        return $this;
    }

    protected function executeDirect(): static
    {
        $this->writer = $this->writer
            ->header($this->headers)
            ->data($this->data);

        if ($this->options->headerStyle) {
            $this->writer = $this->writer->setRow('A1', 18, $this->style());
        }

        return $this;
    }

    protected function executeCustom(): static
    {
        $startRow = $this->options->row - 1;
        $startColumn = Column::labelToNumber($this->options->col);
        $column = $startColumn - 1;
        foreach ($this->headers as $value) {
            $this->setCellValue(++$column, $startRow, $value);
        }

        if ($this->options->headerStyle) {
            $this->writer = $this->writer->setRow($this->options->col.$this->options->row, 18, $this->style());
        }

        ++$startRow;
        foreach ($this->data as $dataRow) {
            $column = $startColumn - 1;
            foreach ($dataRow as $value) {
                $this->setCellValue(++$column, $startRow, $value);
            }
            ++$startRow;
        }

        return $this;
    }

    protected function createExcel(): void
    {
        if (!\extension_loaded('xlswriter')) {
            throw new WriterException('No load extension xlswriter');
        }

        $options = [];
        $options['path'] = $this->options->path ?? sys_get_temp_dir();

        if (isset($options['memory']) && true === $options['memory']) {
            $this->writer = (new Excel($options))
                ->constMemory($this->generateName(), 'DATA');
        } else {
            $this->writer = (new Excel($options))
                ->fileName($this->generateName(), 'DATA');
        }
    }

    protected function generateName(): string
    {
        return uniqid();
    }

    protected function style()
    {
        $fileHandle = $this->writer->getHandle();

        return (new Format($fileHandle))
            ->bold()
            ->align(Format::FORMAT_ALIGN_JUSTIFY, Format::FORMAT_ALIGN_VERTICAL_CENTER)
//            ->background(Format::COLOR_SILVER)
            ->toResource();
    }
}
