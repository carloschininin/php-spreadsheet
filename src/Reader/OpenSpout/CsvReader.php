<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader\OpenSpout;

use CarlosChininin\Spreadsheet\Reader\ReaderOptions;
use OpenSpout\Reader\CSV\Options;
use OpenSpout\Reader\CSV\Reader;

class CsvReader
{
    public const DELIMITER = ';';

    public static function create(ReaderOptions $defaultOptions): Reader
    {
        $options = new Options();
        $options->FIELD_DELIMITER = $defaultOptions->fieldDelimiter ?? self::DELIMITER;
        $options->SHOULD_PRESERVE_EMPTY_ROWS = true;
        // $options->FIELD_ENCLOSURE = '@';

        return new Reader($options);
    }
}
