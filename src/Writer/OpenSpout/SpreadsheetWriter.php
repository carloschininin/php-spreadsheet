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
use OpenSpout\Common\Entity\Style\BorderName;
use OpenSpout\Common\Entity\Style\BorderPart;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\CellVerticalAlignment;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Writer\AbstractWriter;
use OpenSpout\Writer\CSV\Writer as CsvWriter;
use OpenSpout\Writer\Exception\InvalidSheetNameException;
use OpenSpout\Writer\Exception\WriterNotOpenedException;
use OpenSpout\Writer\ODS\Writer as OdsWriter;
use OpenSpout\Writer\XLSX\Options as XlsxOptions;
use OpenSpout\Writer\XLSX\Properties as XlsxProperties;
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

    public function __construct(
        protected readonly iterable $data = [],
        protected readonly iterable $headers = [],
        protected readonly WriterOptions $options = new WriterOptions(),
    ) {
        $this->writer = $this->createWriter();
        $this->filePath = $this->createTempFilePath();
    }

    public function __destruct()
    {
        if (WriterState::Open === $this->state) {
            $this->closeWriter();
        }
    }

    public function execute(bool $direct = true): static
    {
        try {
            $this->openWriter();
            $headerStyle = $this->options->headerStyle ? $this->headerStyle() : null;
            $this->writer->addRow(Helper::createRowHeader($this->headers, $headerStyle));
            foreach ($this->data as $dataRow) {
                $this->writer->addRow(Helper::createRow($dataRow, $this->options));
            }
            $this->closeWriter();
        } catch (IOException|WriterNotOpenedException|InvalidSheetNameException $e) {
            $this->closeWriter();
            throw new WriterException('Failed create spreadsheet: '.$e->getMessage());
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
        $this->closeWriter();
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
            $this->openWriter();
            $rowStyle = $style instanceof Style ? $style : null;
            foreach ($data as $dataRow) {
                $this->writer->addRow(
                    null !== $rowStyle
                        ? Helper::createRowHeader($dataRow, $rowStyle)
                        : Helper::createRow($dataRow, $this->options),
                );
            }
        } catch (IOException|WriterNotOpenedException|InvalidSheetNameException $e) {
            $this->closeWriter();
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
        $border = new Border(
            new BorderPart(BorderName::TOP),
            new BorderPart(BorderName::BOTTOM),
            new BorderPart(BorderName::LEFT),
            new BorderPart(BorderName::RIGHT),
        );

        return new Style(
            fontBold: true,
            fontSize: 11,
            fontColor: Color::WHITE->value,
            cellAlignment: CellAlignment::CENTER,
            cellVerticalAlignment: CellVerticalAlignment::CENTER,
            shouldWrapText: true,
            border: $border,
            backgroundColor: Color::HEADER->value,
        );
    }

    public function addSheet(string $title, bool $isActive = true): static
    {
        if (!method_exists($this->writer, 'addNewSheetAndMakeItCurrent')) {
            return $this;
        }

        try {
            $this->openWriter();
            $sheet = $this->writer->addNewSheetAndMakeItCurrent();
            $sheet->setName($this->normalizeSheetTitle($title));
        } catch (IOException|WriterNotOpenedException|InvalidSheetNameException $e) {
            $this->closeWriter();
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

    private function createWriter(): AbstractWriter
    {
        $writer = match ($this->options->type) {
            SpreadsheetType::ODS => new OdsWriter(),
            SpreadsheetType::CSV => new CsvWriter(),
            default => new XlsxWriter(new XlsxOptions(properties: new XlsxProperties(
                application: 'PIDIA SRL',
                creator: 'PIDIA SRL',
                lastModifiedBy: 'PIDIA SRL',
            ))),
        };

        if (!$writer instanceof XlsxWriter) {
            $writer->setCreator('PIDIA SRL');
        }

        return $writer;
    }

    private function createTempFilePath(): string
    {
        $filePath = tempnam($this->options->path ?? sys_get_temp_dir(), uniqid());
        if (false === $filePath) {
            throw new WriterException('Failed create temporary file');
        }

        return $filePath;
    }

    /**
     * @throws IOException
     * @throws InvalidSheetNameException
     */
    private function openWriter(): void
    {
        if (WriterState::Open === $this->state) {
            return;
        }

        $this->writer->openToFile($this->filePath);
        $this->state = WriterState::Open;
        $this->renameInitialSheet();
    }

    private function closeWriter(): void
    {
        if (WriterState::Close === $this->state) {
            return;
        }

        $this->writer->close();
        $this->state = WriterState::Close;
    }

    /**
     * @throws InvalidSheetNameException
     * @throws WriterNotOpenedException
     */
    private function renameInitialSheet(): void
    {
        if (!$this->options->nameSheet || !method_exists($this->writer, 'getCurrentSheet')) {
            return;
        }

        $this->writer->getCurrentSheet()->setName($this->normalizeSheetTitle($this->options->nameSheet));
    }

    private function normalizeSheetTitle(string $title): string
    {
        return mb_substr($title, 0, 31);
    }
}
