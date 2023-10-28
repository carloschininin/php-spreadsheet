<?php
declare(strict_types=1);

namespace CarlosChininin\Spreadsheet\Shared;

enum SpreadsheetType: string
{
    case XLSX = 'xlsx';
    case XLS = 'xls';
    case CSV = 'csv';
    case ODS = 'ods';
}
