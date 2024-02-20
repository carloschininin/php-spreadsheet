<?php
declare(strict_types=1);

namespace CarlosChininin\Spreadsheet\Reader\OpenSpout;

use CarlosChininin\Spreadsheet\Reader\ReaderOptions;
use OpenSpout\Reader\XLSX\Options;
use OpenSpout\Reader\XLSX\Reader;

class XlsxReader
{
    public static function create(ReaderOptions $options): Reader
    {
        $options = new Options();
//        $options->SHOULD_FORMAT_DATES = true;

        return new Reader($options);
    }
}