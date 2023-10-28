<?php
declare(strict_types=1);

namespace CarlosChininin\Spreadsheet\Shared;

enum DataFormat: string
{
    case GENERAL = 'General';

    case TEXT = '@';

    case NUMBER = '0';
    case NUMBER_00 = '0.00';
    case NUMBER_COMMA_SEPARATED1 = '#,##0.00';
    case NUMBER_COMMA_SEPARATED2 = '#,##0.00_-';

    case PERCENTAGE = '0%';
    case PERCENTAGE_00 = '0.00%';

    case DATE_YYYYMMDD = 'yyyy-mm-dd';
    case DATE_DDMMYYYY = 'dd/mm/yyyy';
    case DATE_DDMMYYYY2 = 'dd-mm-yyyy';
    case DATE_DMYSLASH = 'd/m/yy';
    case DATE_DMYMINUS = 'd-m-yy';
    case DATE_DMMINUS = 'd-m';
    case DATE_MYMINUS = 'm-yy';
    case DATE_XLSX14 = 'mm-dd-yy';
    case DATE_XLSX15 = 'd-mmm-yy';
    case DATE_XLSX16 = 'd-mmm';
    case DATE_XLSX17 = 'mmm-yy';
    case DATE_XLSX22 = 'm/d/yy h:mm';
    case DATE_DATETIME = 'd/m/yy h:mm';
    case DATE_TIME1 = 'h:mm AM/PM';
    case DATE_TIME2 = 'h:mm:ss AM/PM';
    case DATE_TIME3 = 'h:mm';
    case DATE_TIME4 = 'h:mm:ss';
    case DATE_TIME5 = 'mm:ss';
    case DATE_TIME6 = 'i:s.S';
    case DATE_TIME7 = 'h:mm:ss;@';
    case DATE_YYYYMMDDSLASH = 'yyyy/mm/dd;@';

    case CURRENCY_SOL_SIMPLE = '"S/"#,##0.00_-';
    case CURRENCY_SOL = 'S/#,##0_-';
    case CURRENCY_USD_SIMPLE = '"$"#,##0.00_-';
    case CURRENCY_USD = '$#,##0_-';
    case CURRENCY_EUR_SIMPLE = '#,##0.00_-"€"';
    case CURRENCY_EUR = '#,##0_-"€"';
    case ACCOUNTING_USD = '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';
}
