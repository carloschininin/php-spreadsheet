<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Shared;

enum DataType: string
{
    case STRING = 's';
    case FORMULA = 'f';
    case NUMERIC = 'n';
    case BOOL = 'b';
    case NULL = 'null';
    case ERROR = 'e';
    case DATE = 'd';
}
