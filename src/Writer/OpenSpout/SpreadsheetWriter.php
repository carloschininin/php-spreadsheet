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
use OpenSpout\Writer\Exception\WriterNotOpenedException;
use OpenSpout\Writer\ODS\Writer as OdsWriter;
use OpenSpout\Writer\XLSX\Writer as XlsxWriter;
use Symfony\Component\HttpFoundation\Response;

/**
 * Spreadsheet writer medium
 */
class SpreadsheetWriter implements WriterInterface
{
    use WriterTrait;

    protected AbstractWriter $writer;
    protected string $filePath;

    public function __construct(
        protected readonly array $data = [],
        protected readonly array $headers = [],
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

    public function setCellValue(string|int $col, int $row, mixed $value, DataFormat $format = null, DataType $type = null): static
    {
        return $this;
    }

    public function download(string $fileName, bool $useZip = false): Response
    {
        $fileName = File::updateFileName($fileName, $this->options->type);

        return $useZip
            ? File::downloadZip($fileName, $this->filePath)
            : File::download($fileName, $this->filePath);
    }

    public function columnAutoSize(string $start = null, string $end = null): static
    {
        // No implement in this library
        return $this;
    }

    protected function headerStyle(): Style
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
}
