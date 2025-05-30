<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Shared;

enum SpreadsheetType: string
{
    case XLSX = 'xlsx';
    case XLS = 'xls';
    case CSV = 'csv';
    case ODS = 'ods';
}
