<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Shared;

final class DataHelper
{
    public static function boolToString(bool $value): string
    {
        return $value ? 'SI' : 'NO';
    }
}
