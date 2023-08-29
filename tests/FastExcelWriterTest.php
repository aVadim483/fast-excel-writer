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

    protected function getValues($cells): array
    {
        $result = [];
        foreach ($cells as $cell) {
            $result[] = $this->getValue($cell);
        }

        return $result;
    }

    protected function getStyle($cell, $flat = false): array
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

    protected function defaultStyle(): array
    {
        return [
            'font-name' => 'Arial',
            'font-charset' => '1',
            'font-family' => '2',
            'font-size' => '10',
            'fill-pattern' => 'none',
            'border-left-style' => null,
            'border-right-style' => null,
            'border-top-style' => null,
            'border-bottom-style' => null,
            'border-diagonal-style' => null,
            'format-num-id' => 164,
            'format-pattern' => 'GENERAL',
        ];

    }


    protected function checkDefaultStyle($style)
    {
        foreach ($this->defaultStyle() as $key => $val) {
            $this->assertEquals($val, $style[$key]);
        }
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
        $sheet->writeRow(['text0 text0 text0 text0', time(), 0.0]);

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

        $this->checkDefaultStyle($this->getStyle('a7', true));

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
        $this->assertEquals('DD.MM.YYYY HH:MM:SS', $style['format-pattern']);

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
        $area->setValue('a2:c2', $title)
            ->applyFontSize(24)
            ->applyFontStyleBold()
        //    ->applyTextCenter()
        ;

        $area
            ->setValue('a4:a5', 'a4:a5')
            ->setValue('b4:c4', 'b4:c4')
            ->setValue('d4', 'd4')
            ->setValue('c5', 'c5')
            ->setValue('b5', 'b5')
        ;
        $area->withRange('a4:d5')
            ->applyBgColor('#ccc')
            ->applyFontStyleBold()
            ->applyOuterBorder('thin')
            ->applyInnerBorder('thick');
        $sheet->endAreas();

        $data = [
            ['A', 'B', 'C', 'D'],
            ['AA', 'BB', 'CC', 'DD'],
            ['AAA', 'BBB', 'CCC', 'DDD'],
        ];
        foreach ($data as $rowData) {
            $sheet->writeRow($rowData);
        }

        $excel->save($testFileName);
        $this->assertTrue(file_exists($testFileName));

        $this->excelReader = ExcelReader::open($testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $value = $this->getValue('a1');
        $this->assertEquals('', (string)$value);
        $this->checkDefaultStyle($this->getStyle('a1', true));

        $this->assertEquals($title, $this->getValue('a2'));

        $style = $this->getStyle('a2', true);
        $this->assertEquals('24', $style['font-size']);
        $this->assertEquals(1, $style['font-style-bold']);

        $this->assertEquals($data[0], $this->getValues(['A6', 'B6', 'C6', 'D6']));
        $this->assertEquals($data[1], $this->getValues(['A7', 'B7', 'C7', 'D7']));
        $this->assertEquals($data[2], $this->getValues(['A8', 'B8', 'C8', 'D8']));

        unlink($testFileName);
    }

    public function testExcelWriter4()
    {
        $testFileName = __DIR__ . '/test4.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create(['Demo']);
        $excel->setLocale('en');
        $sheet = $excel->getSheet();
        $sheet->setColFormats(['C' => 0, 'D' => '@money', 'E' => '0.00', 'F' => '@']);

        $area = $sheet->beginArea('c3');
        $area->writeRow([1, 1, '=RC[-1]+RC[-2]', 1]);
        $area->writeRow([2, 2, '=RC[-1]+RC[-2]', 2]);
        $area->writeRow([3, 3, '=RC[-1]+RC[-2]', 3]);

        $sheet->endAreas();
        $excel->save($testFileName);
        $this->assertTrue(file_exists($testFileName));

        $this->excelReader = ExcelReader::open($testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $this->assertEquals([1, 1, '=D3+C3', '1'], $this->getValues(['c3', 'd3', 'e3', 'f3']));

        unlink($testFileName);
    }

    public function testExcelWriter5()
    {
        $testFileName = __DIR__ . '/test5.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create(['Demo']);
        $excel->setLocale('en');
        $sheet = $excel->getSheet();
        $sheet->setTopLeftCell('c3');

        $sheet->writeRow([1, 11, 111]);
        $sheet->writeRow([2, 22, 222]);
        $sheet->writeCell(3);
        $sheet->writeCell(33);
        $sheet->writeCell(333);

        $excel->save($testFileName);
        $this->assertTrue(file_exists($testFileName));

        $this->excelReader = ExcelReader::open($testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $this->assertEquals([1, 11, 111], $this->getValues(['c3', 'd3', 'e3']));
        $this->assertEquals([2, 22, 222], $this->getValues(['c4', 'd4', 'e4']));
        $this->assertEquals([3, 33, 333], $this->getValues(['c5', 'd5', 'e5']));

        unlink($testFileName);
    }

    public function testExcelWriterNotesAndImages()
    {
        $testFileName = __DIR__ . '/test_notes_images.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create(['Demo']);
        $sheet = $excel->getSheet();

        $sheet->writeCell('Text to A1');
        $sheet->addNote('A1', 'This is a note for cell A1');

        $sheet->addNote('b2', 'This is a note for cell B2', ['width' => '200pt', 'fill_color' => 'f99']);
        $sheet->addNote('c3', 'This is a note for cell C3', ['width' => '200pt', 'height' => '100pt']);
        $sheet->addNote('D4', 'Note for D4', ['width' => 200, 'height' => 300, 'fill_color' => '#FEDCBA', 'show' => true]);

        $sheet->writeCell('Text to d4')->addNote('This is a note for D4');
        $sheet->writeTo('e5', 'Text to E5')->addNote('Note for C1', ['width' => '200pt', 'height' => '100pt']);

        $sheet->addNote('E4:F8', 'This note will added to E4');

        $sheet->addImage('A10', __DIR__ . '/../demo/logo/excel-logo.gif');
        $sheet->addImage('B10', __DIR__ . '/../demo/logo/excel-logo.jpg');
        $sheet->addImage('C10', __DIR__ . '/../demo/logo/excel-logo.png');
        $sheet->addImage('D10', __DIR__ . '/../demo/logo/excel-logo.svg');
        $sheet->addImage('E10', __DIR__ . '/../demo/logo/excel-logo.webp');

        $excel->save($testFileName);
        $this->assertTrue(file_exists($testFileName));

        $this->excelReader = ExcelReader::open($testFileName);
        $sheet = $this->excelReader->sheet();
        $this->assertEquals(5, $sheet->countImages());
        $testList = [
            'A10' => [
                'image_name' => 'excel-logo.gif',
                'file_name' => 'image1.gif',
            ],
            'B10' => [
                'image_name' => 'excel-logo.jpg',
                'file_name' => 'image2.jpg',
            ],
            'C10' => [
                'image_name' => 'excel-logo.png',
                'file_name' => 'image3.png',
            ],
            'D10' => [
                'image_name' => 'excel-logo.svg',
                'file_name' => 'image4.svg',
            ],
            'E10' => [
                'image_name' => 'excel-logo.webp',
                'file_name' => 'image5.webp',
            ],
        ];
        $this->assertEquals($testList, $sheet->getImageList());

        unlink($testFileName);
    }

}
