<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer;

use CarlosChininin\Spreadsheet\Shared\DataFormat;
use CarlosChininin\Spreadsheet\Shared\DataType;
use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;
use Symfony\Component\HttpFoundation\Response;

/**
 * Generic spreadsheet writer with automatic backend fallback.
 */
class SpreadsheetWriter implements WriterInterface
{
    use WriterTrait;

    private ?WriterInterface $writer = null;
    private ?WriterBackend $backend = null;
    private bool $executed = false;
    private bool $hasManualWrites = false;

    public function __construct(
        protected readonly iterable $data = [],
        protected readonly iterable $headers = [],
        protected readonly WriterOptions $options = new WriterOptions(),
    ) {
    }

    public function execute(bool $direct = true): static
    {
        $this->resolveWriter()->execute($direct);
        $this->executed = true;

        return $this;
    }

    public function download(string $fileName, bool $useZip = false): Response
    {
        $writer = $this->resolveWriter();

        if (!$this->executed && !$this->hasManualWrites) {
            $writer->execute();
            $this->executed = true;
        }

        return $writer->download($fileName, $useZip);
    }

    public function backend(): ?WriterBackend
    {
        return $this->backend;
    }

    public function setCellValue(string|int $col, int $row, mixed $value, DataFormat|string|null $format = null, ?DataType $type = null): static
    {
        $this->resolveCellWriter('setCellValue')->setCellValue($col, $row, $value, $format, $type);
        $this->hasManualWrites = true;

        return $this;
    }

    public function columnAutoSize(string|int|null $start = null, string|int|null $end = null): static
    {
        $this->resolvePhpSpreadsheetWriter('columnAutoSize')->columnAutoSize($start, $end);

        return $this;
    }

    public function mergeCells(string|array $start, string|array $end, mixed $value = null, ?array $style = null): static
    {
        $this->resolvePhpSpreadsheetWriter('mergeCells')->mergeCells($start, $end, $value, $style);
        $this->hasManualWrites = $this->hasManualWrites || null !== $value;

        return $this;
    }

    public function styleCells(string|array $start, string|array $end, array $style): static
    {
        $this->resolvePhpSpreadsheetWriter('styleCells')->styleCells($start, $end, $style);

        return $this;
    }

    public function fromArray(string|int $col, int $row, mixed $data, mixed $style = null): static
    {
        $writer = null === $style
            ? $this->resolveTabularWriter('fromArray')
            : $this->resolvePhpSpreadsheetWriter('fromArray with style');

        $writer->fromArray($col, $row, $data, $style);
        $this->hasManualWrites = true;

        return $this;
    }

    public function formatCells(DataFormat|string $format, string|array $start, string|array|null $end = null): static
    {
        $this->resolvePhpSpreadsheetWriter('formatCells')->formatCells($format, $start, $end);

        return $this;
    }

    public function addSheet(string $title, bool $isActive = true): static
    {
        $this->resolveTabularWriter('addSheet')->addSheet($title, $isActive);

        return $this;
    }

    public function activeSheet(string $title): static
    {
        $this->resolvePhpSpreadsheetWriter('activeSheet')->activeSheet($title);

        return $this;
    }

    public function removeInitialSheet(): static
    {
        $this->resolvePhpSpreadsheetWriter('removeInitialSheet')->removeInitialSheet();

        return $this;
    }

    public function renameSheet(int|string $sheetIndexOrTitle, string $newTitle): bool
    {
        return $this->resolvePhpSpreadsheetWriter('renameSheet')->renameSheet($sheetIndexOrTitle, $newTitle);
    }

    private function resolveWriter(): WriterInterface
    {
        if (null !== $this->writer) {
            return $this->writer;
        }

        if ($this->canUseExcelWriter()) {
            return $this->createExcelWriter();
        }

        if ($this->canUseOpenSpoutWriter()) {
            return $this->createOpenSpoutWriter();
        }

        return $this->createPhpSpreadsheetWriter();
    }

    private function resolveCellWriter(string $operation): WriterInterface
    {
        if (null !== $this->writer) {
            $this->assertBackendSupports($operation, WriterBackend::ExcelWriter, WriterBackend::PhpSpreadsheet);

            return $this->writer;
        }

        return $this->canUseExcelWriter()
            ? $this->createExcelWriter()
            : $this->createPhpSpreadsheetWriter();
    }

    private function resolveTabularWriter(string $operation): WriterInterface
    {
        if (null !== $this->writer) {
            $this->assertBackendSupports($operation, WriterBackend::OpenSpout, WriterBackend::PhpSpreadsheet);

            return $this->writer;
        }

        return $this->canUseOpenSpoutWriter()
            ? $this->createOpenSpoutWriter()
            : $this->createPhpSpreadsheetWriter();
    }

    private function resolvePhpSpreadsheetWriter(string $operation): WriterInterface
    {
        if (null !== $this->writer) {
            $this->assertBackendSupports($operation, WriterBackend::PhpSpreadsheet);

            return $this->writer;
        }

        return $this->createPhpSpreadsheetWriter();
    }

    private function assertBackendSupports(string $operation, WriterBackend ...$supportedBackends): void
    {
        if (null !== $this->backend && \in_array($this->backend, $supportedBackends, true)) {
            return;
        }

        throw new WriterException(\sprintf('Operation "%s" is not supported by selected %s backend', $operation, $this->backend?->value ?? 'unknown'));
    }

    private function canUseExcelWriter(): bool
    {
        return SpreadsheetType::XLSX === $this->options->type
            && \extension_loaded('xlswriter')
            && class_exists(\Vtiful\Kernel\Excel::class);
    }

    private function canUseOpenSpoutWriter(): bool
    {
        return \in_array($this->options->type, [SpreadsheetType::XLSX, SpreadsheetType::CSV, SpreadsheetType::ODS], true)
            && class_exists(\OpenSpout\Writer\XLSX\Writer::class);
    }

    private function createExcelWriter(): WriterInterface
    {
        $this->backend = WriterBackend::ExcelWriter;
        $this->writer = new ExcelWriter\SpreadsheetWriter($this->data, $this->headers, $this->options);

        return $this->writer;
    }

    private function createOpenSpoutWriter(): WriterInterface
    {
        $this->backend = WriterBackend::OpenSpout;
        $this->writer = new OpenSpout\SpreadsheetWriter($this->data, $this->headers, $this->options);

        return $this->writer;
    }

    private function createPhpSpreadsheetWriter(): WriterInterface
    {
        $this->backend = WriterBackend::PhpSpreadsheet;
        $this->writer = new PhpSpreadsheet\SpreadsheetWriter($this->data, $this->headers, $this->options);

        return $this->writer;
    }
}
