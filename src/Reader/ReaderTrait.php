<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader;

use CarlosChininin\Spreadsheet\Shared\SpreadsheetType;

trait ReaderTrait
{
    protected SpreadsheetType $type;
    protected ReaderOptions $options;
    protected string $fileName;

    protected function getType(string $fileName): SpreadsheetType
    {
        $extension = pathinfo($fileName, PATHINFO_EXTENSION);
        foreach (SpreadsheetType::cases() as $type) {
            if ($type->value === $extension) {
                return $type;
            }
        }

        throw new ReaderException(sprintf('File %s not compatible', $fileName));
    }

    protected function validaType(string $fileName, SpreadsheetType $type): void
    {
        if (!str_ends_with($fileName, $type->value)) {
            throw new ReaderException(sprintf('%s file not compatible with %s', $fileName, $type->value));
        }
    }
}
