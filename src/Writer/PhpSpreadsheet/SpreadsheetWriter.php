<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer\PhpSpreadsheet;

use CarlosChininin\Spreadsheet\Shared\Color;
use CarlosChininin\Spreadsheet\Shared\DataFormat;
use CarlosChininin\Spreadsheet\Shared\DataHelper;
use CarlosChininin\Spreadsheet\Shared\DataType;
use CarlosChininin\Spreadsheet\Shared\File;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use CarlosChininin\Spreadsheet\Writer\WriterException;
use CarlosChininin\Spreadsheet\Writer\WriterInterface;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;
use CarlosChininin\Spreadsheet\Writer\WriterTrait;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType as PhpSpreadsheetDataType;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
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
        $this->filePath = $this->createTempFilePath();
        $this->writer = new Spreadsheet();
        $this->processOptions($this->options);
    }

    public function __destruct()
    {
        if (isset($this->writer)) {
            $this->writer->disconnectWorksheets();
        }
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
            $this->writeDataRows($row + 1, $this->columnIndex($col));
        }

        if ($this->options->headerStyle) {
            $this->styleCells($this->col().$this->row(), $this->endCol().$this->row(), $this->headerStyle());
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

        if (\is_bool($value) && null === $type) {
            $sheet->setCellValue($position, DataHelper::boolToString($value));

            return $this;
        }

        if (null === $type) {
            $sheet->setCellValue($position, $value);

            return $this;
        }

        [$value, $dataType] = $this->normalizeTypedValue($value, $type);
        $sheet->setCellValueExplicit($position, $value, $dataType);

        return $this;
    }

    public function fromArray(string|int $col, int $row, mixed $data, mixed $style = null): static
    {
        $colLabel = $this->columnLabel($col);

        $sheet = $this->writer->getActiveSheet();
        $sheet->fromArray($data, startCell: $colLabel.$row, strictNullComparison: true);

        if (null !== $style) {
            try {
                $sheet->getStyle($this->arrayRange($col, $row, $data))->applyFromArray($style);
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
        $columnStart = $this->columnIndex($start ?? $this->col());
        $columnEnd = $this->columnIndex($end ?? $this->endCol());

        for ($i = $columnStart; $i <= $columnEnd; ++$i) {
            $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($i))->setAutoSize(true);
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
        $writer = $this->createWriter();

        try {
            $writer->save($this->filePath);
        } catch (Exception) {
            throw new WriterException('failed save file');
        }
    }

    public function addSheet(string $title, bool $isActive = true): static
    {
        try {
            $this->writer->addSheet(new Worksheet($this->writer, $title));
        } catch (Exception) {
            throw new WriterException('Failed add sheet');
        }

        if ($isActive) {
            $this->activeSheet($title);
        }

        return $this;
    }

    public function activeSheet(string $title): static
    {
        try {
            $this->writer->setActiveSheetIndexByName($title);
        } catch (Exception) {
            throw new WriterException('Failed active sheet');
        }

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
            $sheet = \is_string($sheetIndexOrTitle)
                ? $this->writer->getSheetByName($sheetIndexOrTitle)
                : $this->writer->getSheet($sheetIndexOrTitle);

            if (null === $sheet) {
                return false;
            }

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

    private function writeDataRows(int $startRow, int $startColumn): void
    {
        foreach ($this->data as $dataRow) {
            $column = $startColumn;
            foreach ($dataRow as $key => $value) {
                $this->setCellValue($column, $startRow, $value, type: $this->options->dataTypes[$key] ?? null);
                ++$column;
            }
            ++$startRow;
        }
    }

    private function normalizeTypedValue(mixed $value, DataType $type): array
    {
        return match ($type) {
            DataType::FLOAT => [(float) $value, PhpSpreadsheetDataType::TYPE_NUMERIC],
            DataType::INT => [(int) $value, PhpSpreadsheetDataType::TYPE_NUMERIC],
            DataType::BOOL => [(bool) $value, PhpSpreadsheetDataType::TYPE_BOOL],
            default => [$value, $this->mapDataType($type)],
        };
    }

    private function mapDataType(DataType $type): string
    {
        return match ($type) {
            DataType::STRING => PhpSpreadsheetDataType::TYPE_STRING,
            DataType::FORMULA => PhpSpreadsheetDataType::TYPE_FORMULA,
            DataType::NUMERIC => PhpSpreadsheetDataType::TYPE_NUMERIC,
            DataType::BOOL => PhpSpreadsheetDataType::TYPE_BOOL,
            DataType::NULL => PhpSpreadsheetDataType::TYPE_NULL,
            DataType::ERROR => PhpSpreadsheetDataType::TYPE_ERROR,
            DataType::DATE => PhpSpreadsheetDataType::TYPE_ISO_DATE,
            DataType::FLOAT, DataType::INT => PhpSpreadsheetDataType::TYPE_NUMERIC,
        };
    }

    private function createWriter(): IWriter
    {
        return match ($this->options->type) {
            SpreadsheetType::CSV => new Csv($this->writer),
            SpreadsheetType::ODS => new Ods($this->writer),
            SpreadsheetType::XLS => new Xls($this->writer),
            default => new Xlsx($this->writer),
        };
    }

    private function createTempFilePath(): string
    {
        $filePath = tempnam($this->options->path ?? sys_get_temp_dir(), uniqid());
        if (false === $filePath) {
            throw new WriterException('Failed create temporary file');
        }

        return $filePath;
    }

    private function columnIndex(string|int $column): int
    {
        return \is_int($column) ? $column : Coordinate::columnIndexFromString($column);
    }

    private function columnLabel(string|int $column): string
    {
        return \is_int($column) ? Coordinate::stringFromColumnIndex($column) : mb_strtoupper($column);
    }

    private function arrayRange(string|int $col, int $row, mixed $data): string
    {
        $firstColumn = $this->columnIndex($col);
        $columnCount = 1;
        if (\is_array($data) && isset($data[0]) && is_countable($data[0])) {
            $columnCount = max(1, \count($data[0]));
        }

        $lastColumn = $firstColumn + $columnCount - 1;

        return Coordinate::stringFromColumnIndex($firstColumn).$row.':'.Coordinate::stringFromColumnIndex($lastColumn).$row;
    }
}
