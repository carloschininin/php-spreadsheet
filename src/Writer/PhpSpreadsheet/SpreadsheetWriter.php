<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer\PhpSpreadsheet;

use CarlosChininin\Spreadsheet\Shared\Color;
use CarlosChininin\Spreadsheet\Shared\Column;
use CarlosChininin\Spreadsheet\Shared\DataFormat;
use CarlosChininin\Spreadsheet\Shared\DataHelper;
use CarlosChininin\Spreadsheet\Shared\DataType;
use CarlosChininin\Spreadsheet\Shared\File;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use CarlosChininin\Spreadsheet\Writer\WriterException;
use CarlosChininin\Spreadsheet\Writer\WriterInterface;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;
use CarlosChininin\Spreadsheet\Writer\WriterTrait;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\HttpFoundation\Response;

/**
 * Spreadsheet writer slow.
 */
class SpreadsheetWriter implements WriterInterface
{
    use WriterTrait;

    protected Spreadsheet $writer;
    protected string $filePath;

    public function __construct(
        protected readonly iterable $data = [],
        protected readonly iterable $headers = [],
        protected readonly WriterOptions $options = new WriterOptions(),
    ) {
        $this->filePath = tempnam($this->options->path ?? sys_get_temp_dir(), uniqid());
        $this->writer = new Spreadsheet();
        $this->processOptions($this->options);
    }

    public function execute(bool $direct = true): static
    {
        $col = $this->col();
        $row = $this->row();
        $sheet = $this->writer->getActiveSheet();
        $sheet->fromArray($this->headers, startCell: $col.$row, strictNullComparison: true);
        if ($direct) {
            $sheet->fromArray($this->data, startCell: $col.($row + 1), strictNullComparison: true);
        } else {
            $startRow = $row + 1;
            $startCol = \ord($this->col()) - \ord('A');
            foreach ($this->data as $dataRow) {
                $column = $startCol;
                foreach ($dataRow as $value) {
                    $this->setCellValue(++$column, $startRow, $value);
                }
                ++$startRow;
            }
        }

        if ($this->options->headerStyle) {
            try {
                $sheet->getStyle($this->range())->applyFromArray($this->headerStyle());
            } catch (Exception) {
            }
        }

        return $this;
    }

    public function setCellValue(string|int $col, int $row, mixed $value, DataFormat|string|null $format = null, ?DataType $type = null): static
    {
        if (null === $value) {
            return $this;
        }

        $format = $format instanceof DataFormat ? $format->value : $format;

        $sheet = $this->writer->getActiveSheet();
        $position = $this->positionCell($col, $row);
        if ($value instanceof \DateTimeInterface) {
            $value = Date::dateTimeToExcel($value);
            $format = $format ?? $this->options->formatDate->value;
            $sheet->setCellValue($position, $value);
            $sheet->getStyle($position)->getNumberFormat()->setFormatCode($format);

            return $this;
        }

        if (\is_float($value)) {
            $format = $format ?? $this->options->formatDecimal->value;
            $sheet->setCellValue($position, $value);
            $sheet->getStyle($position)->getNumberFormat()->setFormatCode($format);

            return $this;
        }

        if (\is_bool($value)) {
            $sheet->setCellValue($position, DataHelper::boolToString($value));

            return $this;
        }

        if (null === $type) {
            $sheet->setCellValue($position, $value);
        } else {
            $sheet->setCellValueExplicit($position, $value, $type->value);
        }

        return $this;
    }

    public function fromArray(string|int $col, int $row, mixed $data, mixed $style = null): static
    {
        $colLabel = \is_string($col) ? $col : Column::numberToLabel($col + \ord('A') - 1);

        $sheet = $this->writer->getActiveSheet();
        $sheet->fromArray($data, startCell: $colLabel.$row, strictNullComparison: true);

        if (null !== $style) {
            $count = isset($data[0]) ? count($data[0]) - 1 : 0;
            $range = $this->positionRange([$col, $row], [$col + $count, $row]);
            try {
                $sheet->getStyle($range)->applyFromArray($style);
            } catch (Exception) {
                throw new WriterException('Failed style cells');
            }
        }

        return $this;
    }

    public function mergeCells(string|array $start, string|array $end, mixed $value = null, ?array $style = null): static
    {
        $range = $this->positionRange($start, $end);
        $sheet = $this->writer->getActiveSheet();
        try {
            $sheet->mergeCells($range);
        } catch (Exception) {
            throw new WriterException('Failed merge cells');
        }

        if (null !== $value) {
            $sheet->setCellValue($start, $value);
        }
        if (null !== $style) {
            try {
                $sheet->getStyle($range)->applyFromArray($style);
            } catch (Exception) {
                throw new WriterException('Failed style cells');
            }
        }

        return $this;
    }

    public function formatCells(DataFormat|string $format, string|array $start, string|array|null $end = null): static
    {
        $format = $format instanceof DataFormat ? $format->value : $format;
        $range = $this->positionRange($start, $end);
        $this->writer->getActiveSheet()
            ->getStyle($range)->getNumberFormat()->setFormatCode($format);

        return $this;
    }

    public function styleCells(string|array $start, string|array $end, array $style): static
    {
        try {
            $range = $this->positionRange($start, $end);
            $this->writer->getActiveSheet()
                ->getStyle($range)
                ->applyFromArray($style);
        } catch (Exception) {
            throw new WriterException('Failed style cells');
        }

        return $this;
    }

    public function download(string $fileName, bool $useZip = false): Response
    {
        $this->saveFile();
        $fileName = File::updateFileName($fileName, $this->options->type);

        return $useZip
            ? File::downloadZip($fileName, $this->filePath)
            : File::download($fileName, $this->filePath);
    }

    public function columnAutoSize(string|int|null $start = null, string|int|null $end = null): static
    {
        $sheet = $this->writer->getActiveSheet();
        if (\is_int($start)) {
            $columnStart = $start + \ord('A') - 1;
        } else {
            $columnStart = $start ? \ord($start) : \ord($this->col());
        }

        if (\is_int($end)) {
            $columnEnd = $end + \ord('A') - 1;
        } else {
            $columnEnd = $end ? \ord($end) : ($this->numCols() + $columnStart + 1);
        }

        for ($i = $columnStart; $i <= $columnEnd; ++$i) {
            $sheet->getColumnDimension(Column::numberToLabel($i))->setAutoSize(true);
        }

        return $this;
    }

    public function headerStyle(): array
    {
        return [
            'font' => [
                'bold' => true,
                'size' => '11',
                'color' => [
                    'rgb' => Color::WHITE->value,
                ],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true, // auto adjusted
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => [
                        'rgb' => Color::BLACK->value,
                    ],
                ],
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => Color::HEADER->value,
                ],
            ],
        ];
    }

    public function saveFile(): void
    {
        $writer = match ($this->options->type) {
            SpreadsheetType::CSV => new Csv($this->writer),
            SpreadsheetType::ODS => new Ods($this->writer),
            SpreadsheetType::XLS => new Xls($this->writer),
            default => new Xlsx($this->writer),
        };

        try {
            $writer->save($this->filePath);
        } catch (Exception) {
            throw new WriterException('failed save file');
        }
    }

    public function addSheet(string $title, bool $isActive = true): static
    {
        $this->writer->addSheet(new Worksheet($this->writer, $title));
        if ($isActive) {
            $this->activeSheet($title);
        }

        return $this;
    }

    public function activeSheet(string $title): static
    {
        $this->writer->setActiveSheetIndexByName($title);

        return $this;
    }

    public function removeInitialSheet(): static
    {
        $this->writer->removeSheetByIndex(0);

        return $this;
    }

    public function renameSheet(int|string $sheetIndexOrTitle, string $newTitle): bool
    {
        if (empty($newTitle) || mb_strlen($newTitle) > 31) {
            throw new \InvalidArgumentException('El título de la hoja debe tener entre 1 y 31 caracteres');
        }

        if (preg_match('/[\\:*?\/\[\]]/', $newTitle)) {
            throw new \InvalidArgumentException('El título contiene caracteres no permitidos: \\ : * ? / [ ]');
        }

        try {
            $sheet = is_string($sheetIndexOrTitle)
                ? $this->writer->getSheetByName($sheetIndexOrTitle)
                : $this->writer->getSheet($sheetIndexOrTitle);

            $sheet->setTitle($newTitle);

            return true;
        } catch (\Exception $e) {
            error_log('Error al renombrar hoja: '.$e->getMessage());

            return false;
        }
    }

    protected function processOptions(WriterOptions $options): void
    {
        if ($options->nameSheet) {
            $this->renameSheet(0, $this->options->nameSheet);
        }

        if ($options->removeSheet) {
            $this->removeInitialSheet();
        }
    }
}
