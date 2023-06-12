<?php

declare(strict_types=1);

namespace avadim\FastExcelWriter;

use PHPUnit\Framework\TestCase;
use avadim\FastExcelReader\Excel as ExcelReader;

final class FastExcelWriterTest extends TestCase
{
    protected ?ExcelReader $excelReader = null;
    protected array $cells = [];


    protected function getValue($cell)
    {
        preg_match('/^(\w+)(\d+)$/', strtoupper($cell), $m);

        return $this->cells[$m[2]][$m[1]]['v'] ?? null;
    }

    protected function getStyle($cell, $flat = false)
    {
        preg_match('/^(\w+)(\d+)$/', strtoupper($cell), $m);
        $styleIdx = $this->cells[$m[2]][$m[1]]['s'] ?? null;
        if ($styleIdx !== null) {
            $style = $this->excelReader->getCompleteStyleByIdx($styleIdx);
            if ($flat) {
                $result = [];
                foreach ($style as $key => $val) {
                    $result = array_merge($result, $val);
                }
            }
            else {
                $result = $style;
            }

            return $result;
        }

        return [];
    }

    public function testExcelWriter1()
    {
        $testFileName = __DIR__ . '/test1.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->getSheet();
        $style = [
            'font' => [
                'font-name' => 'Arial',
                'font-charset' => '1',
                'font-family' => '2',
                'font-size' => '10',
                'font-style-bold' => 1,
            ],
            'format' => [
                'format-align-horizontal' => 'center',
                'format-align-vertical' => 'center',
                'format-text-wrap' => 1,
            ],
            'border' => [
                'border-style' => 'thin'
            ]
        ];

        $data = [
            ['text1 text1 text1 text1', time(), 100.0],
            ['text2 text2 text2 text2', time(), 200.0],
            ['text3 text3 text3 text3', time(), 300.0],
        ];
        foreach ($data as $row) {
            $sheet->writeRow($row, $style);
        }
        foreach ($data as $row) {
            $sheet->writeRow($row)
                ->applyFontStyleBold()
                ->applyTextAlign('center', 'center')
                ->applyBorder(Style::BORDER_STYLE_THIN)
                ->applyRowHeight(24)
            ;
        }

        $excel->save($testFileName);

        $this->assertTrue(file_exists($testFileName));

        $this->excelReader = ExcelReader::open($testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $style = $this->getStyle('A1');
        $this->assertEquals(1, $style['font']['font-style-bold']);

        $this->assertEquals('center', $style['format']['format-align-horizontal']);
        $this->assertEquals('center', $style['format']['format-align-vertical']);

        $this->assertEquals('thin', $style['border']['border-left-style']);
        $this->assertEquals('#000000', $style['border']['border-left-color']);

        $style = $this->getStyle('B5');
        $this->assertEquals(1, $style['font']['font-style-bold']);

        $this->assertEquals('center', $style['format']['format-align-horizontal']);
        $this->assertEquals('center', $style['format']['format-align-vertical']);

        $this->assertEquals('thin', $style['border']['border-left-style']);
        $this->assertEquals('#000000', $style['border']['border-left-color']);

        unlink($testFileName);

        $this->excelReader = null;
        $this->cells = [];
    }

    protected function makeTestFile2($testFileName)
    {
        // PREPARE DEMO DATA
        $lorem = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua';
        $loremLen = strlen($lorem);

        $row = 0;
        $data = [
            [++$row, 'John', 'Smith', date('Y-m-d', 1476356498), '=ROUNDDOWN((TODAY()-RC[-1])/365,0)', 100.0, 1, '=RC[-1]*RC[-2]', 0.15, '=RC[-1]*RC[-2]', substr($lorem, 0, random_int(11, $loremLen))],
            [++$row, 'Giovanni', 'Lee', date('Y-m-d', 1383356321), '=ROUNDDOWN((TODAY()-RC[-1])/365,0)', 200.0, 2, '=RC[-1]*RC[-2]', 0.17, '=RC[-1]*RC[-2]', substr($lorem, 0, random_int(11, $loremLen))],
            [++$row, 'Peter', 'Silva', date('Y-m-d', 1086050000), '=ROUNDDOWN((TODAY()-RC[-1])/365,0)', 300.0, 3, '=RC[-1]*RC[-2]', 0.19, '=RC[-1]*RC[-2]', substr($lorem, 0, random_int(11, $loremLen))],
        ];
        $colors = ['cc9', 'c9c', '9cc'];
        $title = 'This is test XLSX-sheet';

        $excel = Excel::create(['Demo']);
        $sheet = $excel->getSheet();

        $sheet
            ->pageOrientationLandscape()  // set page orientation
            ->pageFitToWidth(1)  // fit width to 1 page
            ->pageFitToHeight(1);// fit height to 1 page

        $headerStyle = [
            'font' => [
                'font-size' => 24,
                'font-style-bold' => 1
            ],
            'format-align-horizontal' => 'center',
            'format-align-vertical' => 'center',
        ];

        $area = $sheet->beginArea();

        $cells = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1'];
        foreach($colors as $n => $color) {
            $cell = $cells[$n];
            // set background colors for specified cells
            $area->setBgColor($cell, $color);
        }

        // Write value to automerged cells
        $area->setValue('A2:K2', $title, $headerStyle);
        $area->setValue('E3:I3', 'avadim/fast-excel-writer', ['hyperlink' => 'https://github.com/aVadim483/fast-excel-writer', 'align'=>'center']);

        $area
            ->setValue('J4', 'Date:', ['text-align' => 'right'])
            ->setValue('K4', date('Y-m-d H:i:s'), ['font-style' => 'bold', 'format' => '@datetime', 'text-align' => 'left'])
        ;


        // Begin new area (specify left top cell)
        $area = $sheet->beginArea('A6');
        $area
            ->setValue('RC:R[1]C', '#') // Merge vertical cells
            ->setValue('RC1:RC2', 'People') // Merge horizontal cells
            ->setValue('R1C1', 'First Name') // Single cell
            ->setValue('R1C2', 'Last Name')
            ->setValue('RC3:R1C3', 'Date')
            ->setValue('RC4:R1C4', 'Age')
            ->setValue('RC5:R1C5', 'Quantity')
            ->setValue('RC6:R1C6', 'Price')
            ->setValue('RC7:R1C7', 'Cost')
            ->setValue('RC8:R1C8', 'Tax Rate')
            ->setValue('RC9:R1C9', 'Tax Value')
            ->setValue('RC10:R1C10', 'Description')
        ;

        $tableHeaderStyle = [
            'font-style' => 'bold',
            'fill-color' => '#eee',
            'text-align' => 'center',
            'vertical-align' => 'center',
            'border-style' => 'thin',
        ];

        $area->setStyle('RC:R1C10', $tableHeaderStyle);

        $area->setOuterBorder('R0C0:R1C10', Style::BORDER_THICK);

        $sheet->writeAreas();

        // Default style options for all next cells
        $sheet->setDefaultStyle(['format-align-vertical' => 'top']);

        // Set widths of columns from the first (A)
        $sheet->setColWidths([5, 16, 16, 'auto']);

        // Set width of the column
        $sheet->setColWidth(['G', 'H', 'J'], 14);

        // Set formats of columns from the first (A); null - default format
        $sheet->setColFormats([null, '@', '@', '@date', '0', '0.00', '@money', '@money']);

        // Set style and width for specified column
        $sheet->setColOptions('K', ['text-wrap' => true, 'width' => 32]);

        // Set options for specified columns in the row
        $cellStyles = ['I' => ['format' => '@percent'], 'j' => ['format' => '@money']];
        foreach($data as $n => $row) {
            if ($n % 2) {
                $rowOptions = ['fill' => '#eee'];
            }
            else {
                $rowOptions = null;
            }
            $sheet->writeRow($row, $rowOptions, $cellStyles);
        }

        $totalRow = [];
        $sheet->writeRow($totalRow, ['font' => 'bold', 'border-top' => 'double']);

        $excel->save($testFileName);
    }

    public function testExcelWriter2()
    {
        $testFileName = __DIR__ . '/test2.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $this->makeTestFile2($testFileName);

        $this->assertTrue(file_exists($testFileName));

        $this->excelReader = ExcelReader::open($testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $style = $this->getStyle('A1');
        $this->assertEquals('#CCCC99', $style['fill']['fill-color']);
        $style = $this->getStyle('B1');
        $this->assertEquals('#CC99CC', $style['fill']['fill-color']);
        $style = $this->getStyle('C1');
        $this->assertEquals('#99CCCC', $style['fill']['fill-color']);

        $value = $this->getValue('A2');
        $this->assertEquals('This is test XLSX-sheet', $value);
        $style = $this->getStyle('A2', true);
        $this->assertEquals(24, (int)$style['font-size']);
        $this->assertEquals(1, (int)$style['font-style-bold']);
        $this->assertEquals('center', $style['format-align-horizontal']);
        $this->assertEquals('center', $style['format-align-vertical']);

        $style = $this->getStyle('K4', true);
        $this->assertEquals(1, (int)$style['font-style-bold']);
        $this->assertEquals('left', $style['format-align-horizontal']);
        $this->assertEquals('DD.MM.YYYY\\ HH:MM:SS', $style['format-pattern']);

        $cells = ['A6', 'B6', 'C6', 'D6', 'E6', 'F6', 'G6', 'H6', 'I6', 'J6', 'K6'];
        foreach ($cells as $cell) {
            $style = $this->getStyle($cell, true);
            $this->assertEquals('thick', $style['border-top-style']);
            $this->assertEquals('#000000', $style['border-top-color']);
            $this->assertEquals('thin', $style['border-bottom-style']);
            $this->assertEquals('#000000', $style['border-bottom-color']);
            $this->assertEquals('solid', $style['fill-pattern']);
            $this->assertEquals('#EEEEEE', $style['fill-color']);
            $this->assertEquals('center', $style['format-align-horizontal']);
            $this->assertEquals('center', $style['format-align-vertical']);
        }

        $style = $this->getStyle('A8', true);
        $this->assertEquals('none', $style['fill-pattern']);
        $this->assertEquals('top', $style['format-align-vertical']);
        $this->assertEquals('GENERAL', $style['format-pattern']);

        $style = $this->getStyle('D8', true);
        $this->assertEquals('none', $style['fill-pattern']);
        $this->assertEquals('top', $style['format-align-vertical']);
        $this->assertEquals('DD.MM.YYYY', $style['format-pattern']);

        $style = $this->getStyle('E8', true);
        $this->assertEquals('none', $style['fill-pattern']);
        $this->assertEquals('top', $style['format-align-vertical']);
        $this->assertEquals('0', $style['format-pattern']);

        $style = $this->getStyle('F9', true);
        $this->assertEquals('solid', $style['fill-pattern']);
        $this->assertEquals('#EEEEEE', $style['fill-color']);
        $this->assertEquals('top', $style['format-align-vertical']);
        $this->assertEquals('0.00', $style['format-pattern']);

        $style = $this->getStyle('K9', true);
        $this->assertEquals('solid', $style['fill-pattern']);
        $this->assertEquals('#EEEEEE', $style['fill-color']);
        $this->assertEquals('top', $style['format-align-vertical']);
        $this->assertEquals('GENERAL', $style['format-pattern']);
        $this->assertEquals(1, $style['format-wrap-text']);

        unlink($testFileName);
    }


    public function testExcelWriter3()
    {
        $testFileName = __DIR__ . '/test3.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create(['Demo']);
        $sheet = $excel->getSheet();

        $area = $sheet->beginArea();
        $title = 'Title';
        $area->setValue('a2:e2', $title)
            ->applyFontSize(24)
            ->applyFontStyleBold();

        $excel->save($testFileName);
        $this->assertTrue(file_exists($testFileName));

        $this->excelReader = ExcelReader::open($testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $this->assertEquals($title, $this->getValue('a2'));

        $style = $this->getStyle('a2', true);
        $this->assertEquals('24', $style['font-size']);
        $this->assertEquals(1, $style['font-style-bold']);

        unlink($testFileName);
    }
}
