<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer\OpenSpout;

use CarlosChininin\Spreadsheet\Shared\DataHelper;
use CarlosChininin\Spreadsheet\Shared\DataType;
use CarlosChininin\Spreadsheet\Writer\WriterOptions;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Cell\DateTimeCell;
use OpenSpout\Common\Entity\Cell\EmptyCell;
use OpenSpout\Common\Entity\Cell\NumericCell;
use OpenSpout\Common\Entity\Cell\StringCell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;

final class Helper
{
    public static function valueToCell(mixed $value, ?WriterOptions $options, int|string $key): Cell
    {
        $dataType = $options->dataTypes[$key] ?? null;
        if (null !== $dataType) {
            $value = self::changeTypeValue($value, $dataType);
        }

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

        if (\is_object($value)) {
            if (\method_exists($value, '__toString')) {
                return new StringCell((string) $value, null);
            }

            return new EmptyCell(null, null);
        }

        return Cell::fromValue($value);
    }

    public static function createRow(array $dataRow, ?WriterOptions $options): Row
    {
        $row = new Row([]);
        foreach ($dataRow as $key => &$value) {
            $row->addCell(self::valueToCell($value, $options, $key));
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

    public static function changeTypeValue(mixed $value, DataType $dataType): mixed
    {
        try {
            return match ($dataType) {
                DataType::NUMERIC, DataType::FLOAT => (float) $value,
                DataType::INT => (int) $value,
                DataType::STRING => (string) $value,
                DataType::DATE => new \DateTime($value),
                DataType::BOOL => (bool) $value,
                DataType::NULL => null,
                default => $value,
            };
        } catch (\Exception) {
        }

        return $value;
    }
}
