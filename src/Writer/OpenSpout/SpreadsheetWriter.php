<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer\OpenSpout;

use CarlosChininin\Spreadsheet\Shared\Color;
use CarlosChininin\Spreadsheet\Shared\DataFormat;
use CarlosChininin\Spreadsheet\Shared\DataType;
use CarlosChininin\Spreadsheet\Shared\File;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use CarlosChininin\Spreadsheet\Writer\WriterException;
use CarlosChininin\Spreadsheet\Writer\WriterInterface;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;
use CarlosChininin\Spreadsheet\Writer\WriterTrait;
use OpenSpout\Common\Entity\Style\Border;
use OpenSpout\Common\Entity\Style\BorderPart;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\CellVerticalAlignment;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\InvalidArgumentException;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Writer\AbstractWriter;
use OpenSpout\Writer\CSV\Writer as CsvWriter;
use OpenSpout\Writer\Exception\InvalidSheetNameException;
use OpenSpout\Writer\Exception\WriterNotOpenedException;
use OpenSpout\Writer\ODS\Writer as OdsWriter;
use OpenSpout\Writer\XLSX\Writer as XlsxWriter;
use Symfony\Component\HttpFoundation\Response;

/**
 * Spreadsheet writer medium.
 */
class SpreadsheetWriter implements WriterInterface
{
    use WriterTrait;

    protected AbstractWriter $writer;
    protected string $filePath;
    protected WriterState $state = WriterState::Close;

    /**
     * @throws \Exception
     */
    public function __construct(
        protected readonly iterable $data = [],
        protected readonly iterable $headers = [],
        protected readonly WriterOptions $options = new WriterOptions(),
    ) {
        $this->writer = match ($this->options->type) {
            SpreadsheetType::ODS => new OdsWriter(),
            SpreadsheetType::CSV => new CsvWriter(),
            default => new XlsxWriter(),
        };
        $this->writer->setCreator('PIDIA SRL');
        $this->filePath = tempnam($this->options->path ?? sys_get_temp_dir(), uniqid());
    }

    public function execute(bool $direct = true): static
    {
        try {
            $this->writer->openToFile($this->filePath);
            // Header
            $headerStyle = $this->options->headerStyle ? $this->headerStyle() : null;
            $this->writer->addRow(Helper::createRowHeader($this->headers, $headerStyle));
            // Data
            $options = $this->options;
            foreach ($this->data as $dataRow) {
                $this->writer->addRow(Helper::createRow($dataRow, $options));
            }

            $this->writer->close();
        } catch (IOException|WriterNotOpenedException) {
            throw new WriterException('Failed create spreadsheet');
        }

        return $this;
    }

    public function setCellValue(string|int $col, int $row, mixed $value, DataFormat|string|null $format = null, ?DataType $type = null): static
    {
        return $this;
    }

    public function mergeCells(string|array $start, string|array $end, mixed $value = null, ?array $style = null): static
    {
        return $this;
    }

    public function styleCells(string|array $start, string|array $end, array $style): static
    {
        return $this;
    }

    public function download(string $fileName, bool $useZip = false): Response
    {
        if (WriterState::Open === $this->state) {
            $this->state = WriterState::Close;
            $this->writer->close();
        }

        $fileName = File::updateFileName($fileName, $this->options->type);

        return $useZip
            ? File::downloadZip($fileName, $this->filePath)
            : File::download($fileName, $this->filePath);
    }

    public function columnAutoSize(string|int|null $start = null, string|int|null $end = null): static
    {
        // No implement in this library
        return $this;
    }

    public function fromArray(string|int $col, int $row, mixed $data, mixed $style = null): static
    {
        try {
            if (WriterState::Close === $this->state) {
                $this->state = WriterState::Open;
                $this->writer->openToFile($this->filePath);
                if ($this->options->nameSheet) {
                    $sheet = $this->writer->getCurrentSheet();
                    $sheet->setName(mb_substr($this->options->nameSheet, 0, 31));
                }
            }

            foreach ($data as $dataRow) {
                $row = $style
                    ? Helper::createRowHeader($dataRow, $style)
                    : Helper::createRow($dataRow, $this->options);

                $this->writer->addRow($row);
            }
        } catch (IOException|WriterNotOpenedException|InvalidSheetNameException $e) {
            throw new WriterException('Failed create spreadsheet: '.$e->getMessage());
        }

        return $this;
    }

    public function formatCells(DataFormat|string $format, string|array $start, string|array|null $end = null): static
    {
        // No implement in this library
        return $this;
    }

    public function headerStyle(): Style
    {
        $style = new Style();
        $style->setFontBold();
        $style->setFontSize(11);
        $style->setFontColor(Color::WHITE->value);
        $style->setShouldWrapText();
        try {
            $style->setCellAlignment(CellAlignment::CENTER);
            $style->setCellVerticalAlignment(CellVerticalAlignment::CENTER);
        } catch (InvalidArgumentException) {
        }
        $style->setBackgroundColor(Color::HEADER->value);
        $border = new Border(
            new BorderPart(Border::TOP),
            new BorderPart(Border::BOTTOM),
            new BorderPart(Border::LEFT),
            new BorderPart(Border::RIGHT),
        );
        $style->setBorder($border);

        return $style;
    }

    public function addSheet(string $title, bool $isActive = true): static
    {
        try {
            if (WriterState::Close === $this->state) {
                $this->state = WriterState::Open;
                $this->writer->openToFile($this->filePath);
            }

            $sheet = $this->writer->addNewSheetAndMakeItCurrent();
            $sheet->setName($title);
        } catch (IOException|WriterNotOpenedException|InvalidSheetNameException $e) {
            throw new WriterException($e->getMessage());
        }

        return $this;
    }

    public function activeSheet(string $title): static
    {
        // No implement in this library
        return $this;
    }

    public function removeInitialSheet(): static
    {
        // No implement in this library
        return $this;
    }

    public function renameSheet(int|string $sheetIndexOrTitle, string $newTitle): bool
    {
        // No implement in this library
        return true;
    }
}
