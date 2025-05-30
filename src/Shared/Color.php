<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Shared;

enum Color: string
{
    case BLACK = '000000';
    case WHITE = 'FFFFFF';
    case RED = 'FF0000';
    case DARK_RED = 'C00000';
    case ORANGE = 'FFC000';
    case YELLOW = 'FFFF00';
    case LIGHT_GREEN = '92D040';
    case GREEN = '00B050';
    case LIGHT_BLUE = '00B0E0';
    case BLUE = '0070C0';
    case DARK_BLUE = '002060';
    case PURPLE = '7030A0';
    case HEADER = '215967';
}
