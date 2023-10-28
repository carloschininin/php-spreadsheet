<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Shared;

final class Column
{
    public static function numberToLabel(int $column): string
    {
        if ($column <= \ord('Z')) {
            return \chr($column);
        }

        $factor = floor(($column - \ord('A')) / 26);
        $base = $factor - 1 + \ord('A');
        $next = $column - 26 * $factor;

        return self::numberToLabel((int) $base).self::numberToLabel((int) $next);
    }

    public static function labelToNumber(string $column): int
    {
        $column = mb_strtoupper($column);
        $length = \mb_strlen($column);

        if (1 === $length) {
            return \ord($column) - \ord('A');
        }
        $lastChar = $column[$length - 1];
        $restOfColumn = mb_substr($column, 0, -1);

        return 26 * self::labelToNumber($restOfColumn) + (\ord($lastChar) - \ord('A') + 1);
    }
}
