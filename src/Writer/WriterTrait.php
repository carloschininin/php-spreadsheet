<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Writer;

use CarlosChininin\Spreadsheet\Shared\Column;

trait WriterTrait
{
    public function col(): string
    {
        return mb_strtoupper($this->options->col);
    }

    public function row(): int
    {
        return $this->options->row;
    }

    public function numCols(): int
    {
        if (null === $this->options->numCols) {
            $this->options->numCols = \count($this->headers);
        }

        return $this->options->numCols;
    }

    public function endCol(): string
    {
        if (null === $this->options->endCol) {
            $this->options->endCol = Column::numberToLabel(\ord($this->col()) + $this->numCols() - 1);
        }

        return $this->options->endCol;
    }

    public function range(): string
    {
        return $this->col().$this->row().':'.$this->endCol().$this->row();
    }
}
