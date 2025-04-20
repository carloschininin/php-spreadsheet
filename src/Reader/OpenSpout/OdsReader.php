<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader\OpenSpout;

use CarlosChininin\Spreadsheet\Reader\ReaderOptions;
use OpenSpout\Reader\ODS\Options;
use OpenSpout\Reader\ODS\Reader;

final class OdsReader
{
    public static function create(?ReaderOptions $options): Reader
    {
        $options = $options ?? new Options();

        return new Reader($options);
    }
}
