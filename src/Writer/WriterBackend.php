<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer;

enum WriterBackend: string
{
    case ExcelWriter = 'excelwriter';
    case OpenSpout = 'openspout';
    case PhpSpreadsheet = 'phpspreadsheet';
}
