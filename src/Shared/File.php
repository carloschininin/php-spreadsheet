<?php

declare(strict_types=1);

/*
 * This file is part of the PIDIA.
 * (c) Carlos Chininin <cio@pidia.pe>
 */

namespace CarlosChininin\Spreadsheet\Shared;

use Symfony\Component\HttpFoundation\BinaryFileResponse;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

final class File
{
    public static function updateFileName(string $fileName, SpreadsheetType $type): string
    {
        if (!str_ends_with(mb_strtolower($fileName), '.'.$type->value)) {
            $fileName .= '.'.$type->value;
        }

        return $fileName;
    }

    public static function download(string $fileName, string $filePath): Response
    {
        $response = new BinaryFileResponse($filePath);
        $response->setContentDisposition(ResponseHeaderBag::DISPOSITION_ATTACHMENT, $fileName);

        return $response;
    }

    public static function downloadZip(string $fileName, string $filePath): Response
    {
        $zip = new \ZipArchive();
        $fileNameZip = $fileName.'.zip';
        $tempFilename = tempnam(sys_get_temp_dir(), $fileNameZip);

        if (true === $zip->open($tempFilename)) {
            $zip->addFile($filePath, $fileName);
            $zip->close();
            unlink($filePath);
        }

        return self::download($fileNameZip, $tempFilename);
    }
}
