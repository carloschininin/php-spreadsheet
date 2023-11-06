<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Reader;

interface ReaderInterface
{
    public function load(string $fileName): static;
    public function data(bool $multipleSheet = false): array;
    public function iterator(callable $callback): static;
}
