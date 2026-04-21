<?php

declare(strict_types=1);

namespace XlsxToJson;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Converts an XLSX workbook (one or more sheets) to a PHP array / JSON.
 *
 * The first row of every sheet is treated as the header row.
 * All subsequent rows become associative arrays keyed by those headers.
 */
class XlsxToJson
{
    /**
     * Read an XLSX file and return an associative array keyed by sheet name.
     * Each value is an array of associative arrays representing the data rows.
     *
     * @param  string  $filePath  Absolute or relative path to the .xlsx file.
     * @return array<string, list<array<string, mixed>>>
     *
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception  When the file cannot be read.
     * @throws \InvalidArgumentException                   When the file does not exist.
     */
    public static function convert(string $filePath): array
    {
        if (!file_exists($filePath)) {
            throw new \InvalidArgumentException("File not found: {$filePath}");
        }

        $spreadsheet = IOFactory::load($filePath);
        $result = [];

        foreach ($spreadsheet->getAllSheets() as $sheet) {
            $result[$sheet->getTitle()] = self::sheetToArray($sheet);
        }

        return $result;
    }

    /**
     * Convert a single worksheet to an array of associative arrays.
     *
     * @return list<array<string, mixed>>
     */
    private static function sheetToArray(Worksheet $sheet): array
    {
        $rows = $sheet->toArray(null, true, true, false);

        if (empty($rows)) {
            return [];
        }

        $headers = array_shift($rows);
        $data = [];

        foreach ($rows as $row) {
            $entry = [];
            foreach ($headers as $index => $header) {
                $entry[(string) $header] = $row[$index] ?? null;
            }
            $data[] = $entry;
        }

        return $data;
    }
}
