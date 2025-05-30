<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader\OpenSpout;

use CarlosChininin\Spreadsheet\Reader\ReaderOptions;
use OpenSpout\Reader\XLSX\Options;
use OpenSpout\Reader\XLSX\Reader;

class XlsxReader
{
    public static function create(?ReaderOptions $options): Reader
    {
        $options = $options ?? new Options();
        //        $options->SHOULD_FORMAT_DATES = true;

        return new Reader($options);
    }
}
