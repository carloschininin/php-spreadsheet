<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer\OpenSpout;

use CarlosChininin\Spreadsheet\Shared\DataHelper;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Cell\DateTimeCell;
use OpenSpout\Common\Entity\Cell\NumericCell;
use OpenSpout\Common\Entity\Cell\StringCell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;

final class Helper
{
    public static function valueToCell(mixed $value, ?WriterOptions $options): Cell
    {
        if ($value instanceof \DateTimeInterface) {
            $style = $options && $options->formatDate ? (new Style())->setFormat($options->formatDate->value) : null;

            return new DateTimeCell($value, $style);
        }

        if (\is_float($value)) {
            $style = $options && $options->formatDecimal ? (new Style())->setFormat($options->formatDecimal->value) : null;

            return new NumericCell($value, $style);
        }

        if (\is_bool($value)) {
            return new StringCell(DataHelper::boolToString($value), null);
        }

        return Cell::fromValue($value);
    }

    public static function createRow(array $dataRow, ?WriterOptions $options): Row
    {
        $row = new Row([]);
        foreach ($dataRow as &$value) {
            $row->addCell(self::valueToCell($value, $options));
        }

        return $row;
    }

    public static function createRowHeader(array $dataRow, Style $rowStyle = null): Row
    {
        $row = new Row([]);
        foreach ($dataRow as &$value) {
            $row->addCell(new StringCell($value, $rowStyle));
        }

        return $row;
    }
}
