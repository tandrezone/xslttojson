<?php

declare(strict_types=1);

namespace XlsxToJson\Tests;

use PHPUnit\Framework\TestCase;
use XlsxToJson\XlsxToJson;

class XlsxToJsonTest extends TestCase
{
    private string $fixtures;

    protected function setUp(): void
    {
        $this->fixtures = __DIR__ . '/fixtures';
    }

    public function testReturnsAnObjectWithAKeyPerSheet(): void
    {
        $result = XlsxToJson::convert($this->fixtures . '/multi-sheet.xlsx');

        $this->assertIsArray($result);
        $this->assertSame(['Employees', 'Offices'], array_keys($result));
    }

    public function testEachSheetValueIsAnArray(): void
    {
        $result = XlsxToJson::convert($this->fixtures . '/multi-sheet.xlsx');

        $this->assertIsArray($result['Employees']);
        $this->assertIsArray($result['Offices']);
    }

    public function testRowsUseFirstRowAsHeaders(): void
    {
        $result = XlsxToJson::convert($this->fixtures . '/multi-sheet.xlsx');
        $firstRow = $result['Employees'][0];

        $this->assertArrayHasKey('Name', $firstRow);
        $this->assertArrayHasKey('Department', $firstRow);
        $this->assertArrayHasKey('Salary', $firstRow);
    }

    public function testRowDataIsCorrectForFirstSheet(): void
    {
        $result = XlsxToJson::convert($this->fixtures . '/multi-sheet.xlsx');

        $this->assertCount(3, $result['Employees']);
        $this->assertSame('Alice', $result['Employees'][0]['Name']);
        $this->assertSame('Engineering', $result['Employees'][0]['Department']);
        $this->assertSame(90000, (int) $result['Employees'][0]['Salary']);
    }

    public function testRowDataIsCorrectForSecondSheet(): void
    {
        $result = XlsxToJson::convert($this->fixtures . '/multi-sheet.xlsx');

        $this->assertCount(2, $result['Offices']);
        $this->assertSame('Paris', $result['Offices'][0]['City']);
        $this->assertSame('London', $result['Offices'][1]['City']);
    }

    public function testWorksWithASingleSheetWorkbook(): void
    {
        $result = XlsxToJson::convert($this->fixtures . '/single-sheet.xlsx');

        $this->assertSame(['Products'], array_keys($result));
        $this->assertCount(2, $result['Products']);
        $this->assertSame('Widget', $result['Products'][0]['Name']);
        $this->assertSame(19.99, (float) $result['Products'][1]['Price']);
    }

    public function testThrowsForMissingFile(): void
    {
        $this->expectException(\InvalidArgumentException::class);
        $this->expectExceptionMessageMatches('/File not found/');

        XlsxToJson::convert($this->fixtures . '/does-not-exist.xlsx');
    }
}
