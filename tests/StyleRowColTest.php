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
    protected array $savedFiles = [];

    protected function setUp(): void
    {
        if (!is_dir($this->tempDir)) {
            mkdir($this->tempDir);
        }
    }

    protected function tearDown(): void
    {
        foreach ($this->savedFiles as $path) {
            if (file_exists($path)) {
                unlink($path);
            }
        }
        $this->savedFiles = [];
    }

    protected function saveCheckRead(Excel $excel, string $filename): ExcelReader
    {
        $path = $this->tempDir . '/' . $filename;
        $this->savedFiles[] = $path;
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
    public function testColHiddenBooleanValues()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        // the 'hidden' attribute is of xs:boolean type: Excel writes 0/1, LibreOffice writes 'false'/'true'
        $sheet->_setColAttributes(0, ['hidden' => 'false', 'width' => 12]);
        $sheet->_setColAttributes(1, ['hidden' => 'true']);
        $sheet->_setColAttributes(2, ['hidden' => '0']);
        $sheet->_setColAttributes(3, ['hidden' => '1']);
        $sheet->setColHidden('E');
        $sheet->setColVisible('F', true);
        $sheet->writeRow(['A1', 'B1', 'C1', 'D1', 'E1', 'F1']);

        $reader = $this->saveCheckRead($excel, 'col_hidden.xlsx');
        $cols = $reader->sheet()->getAllColAttributes();

        $this->assertArrayNotHasKey('hidden', $cols['A'] ?? []);
        $this->assertEquals('1', $cols['B']['hidden'] ?? null);
        $this->assertArrayNotHasKey('hidden', $cols['C'] ?? []);
        $this->assertEquals('1', $cols['D']['hidden'] ?? null);
        $this->assertEquals('1', $cols['E']['hidden'] ?? null);
        $this->assertArrayNotHasKey('hidden', $cols['F'] ?? []);
    }

    public function testRowHiddenBooleanValues()
    {
        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();

        $setRowSettings = new \ReflectionMethod(Sheet::class, '_setRowSettings');
        $setRowSettings->setAccessible(true);

        $sheet->setRowHidden(2);
        $sheet->setRowVisible(3, true);
        // the same xs:boolean values as in a template made by LibreOffice
        $setRowSettings->invoke($sheet, 4, 'hidden', 'false');
        $setRowSettings->invoke($sheet, 5, 'hidden', 'true');

        for ($i = 1; $i <= 5; $i++) {
            $sheet->writeRow(['A' . $i]);
        }

        $reader = $this->saveCheckRead($excel, 'row_hidden.xlsx');
        $rows = $reader->sheet()->getAllRowAttributes();

        $this->assertArrayNotHasKey('hidden', $rows[1]);
        $this->assertEquals('1', $rows[2]['hidden'] ?? null);
        $this->assertEquals('0', $rows[3]['hidden'] ?? null);
        $this->assertEquals('0', $rows[4]['hidden'] ?? null);
        $this->assertEquals('1', $rows[5]['hidden'] ?? null);
    }
}
