<?php

declare(strict_types=1);

namespace avadim\FastExcelWriter;

use avadim\FastExcelReader\Excel as ExcelReader;
use avadim\FastExcelWriter\Style\Style;
use PHPUnit\Framework\TestCase;

final class StyleRowColTest extends TestCase
{
    protected string $tempDir = __DIR__ . '/tmp';
    protected array $cells = [];

    protected function setUp(): void
    {
        if (!is_dir($this->tempDir)) {
            mkdir($this->tempDir);
        }
    }

    protected function saveCheckRead(Excel $excel, string $filename): ExcelReader
    {
        $path = $this->tempDir . '/' . $filename;
        $excel->save($path);
        $this->assertFileExists($path);
        return ExcelReader::open($path);
    }

    protected function getCompleteStyle(ExcelReader $reader, $cellAddress)
    {
        $cells = $reader->readCellsWithStyles();
        if (isset($cells[$cellAddress]['s'])) {
            return $cells[$cellAddress]['s'];
        }
        return [];
    }

    public function testSetRowStyle()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $sheet->writeRow(['A1']);
        $sheet->writeRow(['A2'], ['bg-color' => '#FF0000']);

        $reader = $this->saveCheckRead($excel, 'row_style.xlsx');
        $styleA2 = $this->getCompleteStyle($reader, 'A2');
        $this->assertEquals('#FF0000', $styleA2['fill']['fill-color'] ?? null);
    }

    public function testSetRowStyleArray()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $sheet->writeRow(['A1']);
        $sheet->writeRow(['A2'], ['bg-color' => '#FF0000']);

        $reader = $this->saveCheckRead($excel, 'row_style_array.xlsx');
        $styleA2 = $this->getCompleteStyle($reader, 'A2');
        $this->assertEquals('#FF0000', $styleA2['fill']['fill-color'] ?? null);
    }

    public function testSetRowDataStyle()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $sheet->writeRow(['A1']);
        $sheet->writeRow(['A2'])->applyBgColor('#FF0000');

        $reader = $this->saveCheckRead($excel, 'row_data_style.xlsx');
        $styleA2 = $this->getCompleteStyle($reader, 'A2');
        $this->assertEquals('#FF0000', $styleA2['fill']['fill-color'] ?? null);
    }

    public function testSetRowDataStyleArray()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $style = new Style();
        $sheet->setRowDataStyleArray([
            2 => $style->setBgColor('#FF0000'),
        ]);
        $sheet->writeRow(['A1']);
        $sheet->writeRow(['A2']);

        $reader = $this->saveCheckRead($excel, 'row_data_style_array.xlsx');
        $styleA2 = $this->getCompleteStyle($reader, 'A2');
        $this->assertEquals('#FF0000', $styleA2['fill']['fill-color'] ?? null);
    }

    public function testSetColStyle()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $sheet->setColStyle('B', ['bg-color' => '#FF0000']);
        $sheet->writeRow(['A1', 'B1']);

        $reader = $this->saveCheckRead($excel, 'col_style.xlsx');
        $styleB1 = $this->getCompleteStyle($reader, 'B1');
        $this->assertEquals('#FF0000', $styleB1['fill']['fill-color'] ?? null);
    }

    public function testSetColStyleArray()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $sheet->setColStyleArray([
            'B' => ['bg-color' => '#FF0000'],
        ]);
        $sheet->writeRow(['A1', 'B1']);

        $reader = $this->saveCheckRead($excel, 'col_style_array.xlsx');
        $styleB1 = $this->getCompleteStyle($reader, 'B1');
        $this->assertEquals('#FF0000', $styleB1['fill']['fill-color'] ?? null);
    }

    public function testSetColDataStyle()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $sheet->setColDataStyle('B', ['bg-color' => '#FF0000']);
        $sheet->writeRow(['A1', 'B1']);

        $reader = $this->saveCheckRead($excel, 'col_data_style.xlsx');
        $styleB1 = $this->getCompleteStyle($reader, 'B1');
        $this->assertEquals('#FF0000', $styleB1['fill']['fill-color'] ?? null);
    }

    public function testSetColDataStyleArray()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $sheet->setColDataStyleArray([
            'B' => ['bg-color' => '#FF0000'],
        ]);
        $sheet->writeRow(['A1', 'B1']);

        $reader = $this->saveCheckRead($excel, 'col_data_style_array.xlsx');
        $styleB1 = $this->getCompleteStyle($reader, 'B1');
        $this->assertEquals('#FF0000', $styleB1['fill']['fill-color'] ?? null);
    }
}
