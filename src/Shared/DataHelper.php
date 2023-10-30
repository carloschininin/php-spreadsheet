<?php
declare(strict_types=1);

namespace CarlosChininin\Spreadsheet\Shared;

final class DataHelper
{
    public static function boolToString(bool $value): string
    {
        return $value ? 'SI' : 'NO';
    }
}