<?php

declare(strict_types=1);

namespace avadim\FastExcelWriter;

use avadim\FastExcelReader\Excel as ExcelReader;
use avadim\FastExcelWriter\Charts\Chart;
use avadim\FastExcelWriter\Charts\Legend;
use avadim\FastExcelWriter\Conditional\Conditional;
use avadim\FastExcelWriter\Exceptions\ExceptionAddress;
use avadim\FastExcelWriter\Style\Style;
use PHPUnit\Framework\TestCase;

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
        preg_match('/^([A-Z]+)(\d+)$/', strtoupper($cell), $m);
        $styleIdx = $this->cells[$m[2]][$m[1]]['s'] ?? 0;
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
            'font-name' => 'Calibri',
            'font-charset' => '1',
            'font-family' => '0',
            'font-size' => '11',
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


    protected function saveCheckRead($excel, $testFileName): ExcelReader
    {
        $excel->save($testFileName);
        $this->assertTrue(file_exists($testFileName));
        $this->assertEquals('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', mime_content_type($testFileName));
        $valid = ExcelReader::validate($testFileName, $errors);
        if ($errors) {
            $text = [];
            foreach ($errors as $err) {
                $text[] = [
                    'error' => $err->message,
                    'file' => $err->file,
                ];
            }
            var_dump($text);
        }
        $this->assertTrue($valid);

        if (substr(PHP_OS, 0, 3) === 'WIN') {
            $vbsChecker = __DIR__ . '/win_check_open_xlsx.vbs';
            if (is_file($vbsChecker)) {
                system("$vbsChecker $testFileName", $result);
                $this->assertEquals(0, $result);
            }
        }

        return ExcelReader::open($testFileName);
    }


    public function testExcelWriter0()
    {
        $testFileName = __DIR__ . '/test0.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $tempDir = __DIR__ . '/tmp';
        Excel::setTempDir($tempDir);
        $excel = Excel::create();
        $sheet = $excel->sheet();

        // write row 1, go to row 2
        $sheet
            ->writeCell('A1')
            ->writeCell('B1')
            ->nextCell() // C1
            ->writeCell(fn($sheet) => $sheet->getCurrentCell()) // D1
            ->nextCell() // E1
            ->nextCell() // F1
            ->writeCell(function($sheet) {
                return $sheet->getCurrentCol() . $sheet->getCurrentRow();
            }) // G1
            ->writeTo('F1', 'F1');
        ;
        // write row 2 go to row 3
        $sheet->writeRow(['A2', 'B2', null, 'D2']);

        $sheet->nextRow(); // go to next row - 4
        $sheet->writeTo('F3', 'F3');
        $sheet->nextRow(); // 4
        $sheet->nextRow(); // 5
        $sheet->writeCell('A5'); // 5
        $sheet->skipRow(3); // skip rows 6, 7 and 8 - go to 9
        $sheet->skipRow(); // skip row 9 - go to 10
        $sheet->writeCell('A10');
        $sheet->writeTo('F10', 'F10');
        $sheet->writeTo('D10', 'D10');
        $sheet->writeTo('B10', 'B10');
        $sheet->writeTo('E10', 'E10');

        // rows 11, 12, 13, 14
        $sheet->writeRows([null, [], [''], ['A14', 'b14']]);

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->assertCount(0, glob($tempDir . '/*.*'));
        $this->cells = $this->excelReader->readCells();

        $this->assertEquals('A1', $this->cells['A1']);
        $this->assertEquals('B1', $this->cells['B1']);
        $this->assertEquals('D1', $this->cells['D1']);
        $this->assertEquals('F1', $this->cells['F1']);
        $this->assertEquals('G1', $this->cells['G1']);

        $this->assertEquals('A2', $this->cells['A2']);
        $this->assertEquals('B2', $this->cells['B2']);
        $this->assertEquals('D2', $this->cells['D2']);

        $this->assertEquals(null, $this->cells['A3']);
        $this->assertEquals('F3', $this->cells['F3']);

        $this->assertEquals('A5', $this->cells['A5']);

        $this->assertEquals('A10', $this->cells['A10']);
        $this->assertEquals('B10', $this->cells['B10']);
        $this->assertEquals('D10', $this->cells['D10']);
        $this->assertEquals('E10', $this->cells['E10']);

        $this->assertFalse(isset($this->cells['A9']));
        $this->assertEquals(null, $this->cells['C10']);
        $this->assertEquals('', $this->cells['A13']);
        $this->assertEquals('b14', $this->cells['B14']);

        unlink($testFileName);

        $this->excelReader = null;
        $this->cells = [];
    }


    public function testExcelWriter1()
    {
        $testFileName = __DIR__ . '/test1.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();
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
            1 => ['text1 text1 text1 text1', time(), 100.0],
            2 => ['text2 text2 text2 text2', time(), 200.0],
            3 => ['text3 text3 text3 text3', time(), 300.0],
        ];
        // row 1 with default styles
        $sheet->writeRow($data[1]);
        // rows 2-3 with specified styles
        $sheet->writeRow($data[2], $style);
        $sheet->writeRow($data[3], $style);

        // rows 4-6
        foreach ($data as $row) {
            $sheet->writeRow($row)
                ->applyFontStyleBold()
                ->applyTextAlign('center', 'center')
                ->applyBorder(Style::BORDER_STYLE_THIN)
                ->applyRowHeight(24)
            ;
        }
        // row 7
        $sheet->writeRow(['text0 text0 text0 text0', time(), 0.0])
            ->applyFont('Times New Roman', 18, 'italic', '#f00');

        // write to C8 and move a pointer to D9;
        $sheet->writeTo('C8', 'C8');
        $sheet->writeCell('D8');

        // write to C8 and move a pointer to D9;
        $sheet->setValue('C9', 'C9');
        $sheet->writeCell('replace C9');

        $sheet->setAutofilter();
        $sheet->addNamedRange('b2:c3', 'b2c3');
        $sheet->setPrintArea('a2:f2,a4:f4')->setPrintTitles('1', 'a:b');
        $sheet->pageHeaderFooter('Header', 'Footer');

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $this->assertEquals('text1 text1 text1 text1', $this->cells['1']['A']['v']);
        $this->assertEquals('C8', $this->cells['8']['C']['v']);
        $this->assertEquals('D8', $this->cells['8']['D']['v']);
        $this->assertEquals('replace C9', $this->cells['9']['C']['v']);
        $this->assertFalse(isset($this->cells['9']['D']['v']));

        $style = $this->getStyle('A2');
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

        $this->checkDefaultStyle($this->getStyle('a1', true));

        $style = $this->getStyle('a7', true);

        $this->assertEquals('Times New Roman', $style['font-name']);
        $this->assertEquals('18', $style['font-size']);
        $this->assertEquals('#FF0000', $style['font-color']);
        $this->assertEquals(1, $style['font-style-italic']);

        unlink($testFileName);

        $this->excelReader = null;
        $this->cells = [];
    }


    protected function makeTestFile2($testFileName): ExcelReader
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
        $sheet = $excel->sheet();

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
        $sheet->setColDataStyle('K', ['text-wrap' => true, 'width' => 32]);

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

        return $this->saveCheckRead($excel, $testFileName);
    }


    public function testStyleArray()
    {
        $testFileName = __DIR__ . '/test2.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $this->excelReader = $this->makeTestFile2($testFileName);
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
        // defines from locale
        //$this->assertEquals('DD.MM.YYYY HH:MM:SS', $style['format-pattern']);

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
        // defines from locale
        //$this->assertEquals('DD.MM.YYYY', $style['format-pattern']);

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

        $defaultStyle = $this->defaultStyle();
        $excel = Excel::create(['Demo'], [Style::FONT => $defaultStyle]);
        $sheet = $excel->sheet();

        $area = $sheet->beginArea();
        $area->setValue('b1', '.');
        $title = 'Title';
        $area->setValue('a2:c2', $title)
            ->applyFontSize(24)
            ->applyFontStyleBold()
            ->applyTextCenter()
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
        $sheet->writeArray($data);

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $value = $this->getValue('a1');
        $this->assertEquals('', (string)$value);
        $style = $this->getStyle('a1', true);
        $this->assertEquals($defaultStyle['font-size'], $style['font-size']);

        $this->checkDefaultStyle($this->getStyle('b1', true));

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
        $excel->setDefaultStyle([Style::FONT => [Style::FONT_NAME => 'Century', Style::FONT_SIZE => 21]]);
        $excel->setLocale('en');
        $sheet = $excel->sheet();
        $sheet->setColFormats(['C' => 0, 'D' => '@money', 'E' => '0.00', 'F' => '@']);

        $data = [
            [2, 2, '=RC[-1]+RC[-2]', 2],
            [3, 3, '=RC[-1]+RC[-2]', 3],
            [4, 4, ['=RC[-1]+RC[-2]', 8], 4], // formula & value
        ];
        $area = $sheet->beginArea('b2');
        $area->moveTo('c3');
        $area->writeRow([1, 1, '=RC[-1]+RC[-2]', 1]);
        $area->writeArrayTo('c4', $data);

        $sheet->endAreas();

        $sheet->writeRow([null, null, 6, 7, '=RC[-2]+R3C4']);

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $this->assertEquals([1, 1, '=D3+C3', '1'], $this->getValues(['c3', 'd3', 'e3', 'f3']));

        // formula & value
        $this->assertEquals(8, $this->cells[6]['E']['v']);
        $this->assertEquals('=D6+C6', $this->cells[6]['E']['f']);
        $this->assertEquals('=C7+$D$3', $this->cells[7]['E']['f']);

        $style = $this->getStyle('c3', true);
        $this->assertEquals('21', $style['font-size']);
        $this->assertEquals('Century', $style['font-name']);

        $style = $this->getStyle('a1', true);
        $this->assertEquals('21', $style['font-size']);
        $this->assertEquals('Century', $style['font-name']);

        unlink($testFileName);
    }


    public function testExcelWriter5()
    {
        $testFileName = __DIR__ . '/test5.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create(['Demo1', 'Demo2']);
        $excel->setDefaultFont([Style::FONT_NAME => 'Century']);

        $excel->setLocale('en');
        $sheet = $excel->sheet();
        $sheet->setTopLeftCell('c3');

        $sheet->writeRow([1, 11, 111])->applyFontStyleBold();
        $sheet->writeRow([2, 22, 222]);
        $sheet->writeCell(3);
        $sheet->writeCell('33');
        $sheet->writeCell(333.3);

        $sheet = $excel->sheet('Demo2');
        $sheet->writeHeader(['AAA', 'BBB', 'CCC'])->applyFontStyleBold();

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);

        $this->cells = $this->excelReader->readRows(false, null, true);

        $this->assertEquals([1, 11, 111], $this->getValues(['c3', 'd3', 'e3']));
        $this->assertEquals([2, 22, 222], $this->getValues(['c4', 'd4', 'e4']));

        $this->assertTrue(3 === $this->getValue('c5'));
        $this->assertTrue('33' === $this->getValue('d5'));
        $this->assertTrue(333.3 === $this->getValue('e5'));

        $cells = $this->excelReader->sheet('Demo2')->readCellsWithStyles();
        $this->assertEquals('Century', $cells['A1']['s']['font']['font-name']);

        unlink($testFileName);
    }


    public function testExcelWriterMergedCells()
    {
        $testFileName = __DIR__ . '/test_merged.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $area = $sheet->beginArea();
        $area->setValue('A1', 'A');
        // Write value to automerged cells
        $area->setValue('A2:D2', 'A2:D2');
        $sheet->writeAreas();

        $sheet->writeCell(11);
        $sheet->writeCell(12);
        $sheet->writeCell(13);
        $sheet->nextRow();
        $sheet->writeCell(21);
        $sheet->writeCell(32);
        $sheet->writeCell(43);
        $sheet->mergeCells('A4:C4');
        $sheet->mergeCells('D3:F3');
        $sheet->mergeCells('A5:A7');
        $className = null;
        try {
            $sheet->mergeCells('B1:B2'); // intersect with A2:D2
        }
        catch (\Throwable $e) {
            $className = get_class($e);
        }
        $this->assertSame(ExceptionAddress::class, $className);

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $mergedCells = $this->excelReader->sheet()->getMergedCells();
        $a = ['A2' => 'A2:D2', 'A4' => 'A4:C4', 'D3' => 'D3:F3', 'A5' => 'A5:A7'];
        $this->assertEquals($a, $mergedCells);

        unlink($testFileName);
    }


    public function testExcelWriterNotesAndImages()
    {
        $testFileName = __DIR__ . '/test_notes_images.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create(['Demo']);
        $sheet = $excel->sheet();

        $sheet->writeCell('Text to A1');
        $sheet->addNote('A1', 'This is a note for cell A1');

        $sheet->addNote('b2', 'This is a note for cell B2', ['width' => '200pt', 'fill_color' => 'f99']);
        $sheet->addNote('c3', 'This is a note for cell C3', ['width' => '200pt', 'height' => '100pt']);
        $sheet->addNote('D4', 'Note for D4', ['width' => 200, 'height' => 300, 'fill_color' => '#FEDCBA', 'show' => true]);

        $sheet->writeCell('Text to d4')->addNote('This is a note for D4');
        $sheet->writeTo('e5', 'Text to E5')->addNote('Note for C1', ['width' => '200pt', 'height' => '100pt']);

        $sheet->addNote('E4:F8', 'This note will added to E4');

        $imgDir = __DIR__ . '/../demo/logo';
        $sheet->addImage('A10', $imgDir . '/excel-logo.gif');
        $sheet->addImage('B10', $imgDir . '/excel-logo.jpg');
        $sheet->addImage('C10', $imgDir . '/excel-logo.png');
        $sheet->addImage('D10', $imgDir . '/excel-logo.svg');
        $sheet->addImage('E10', $imgDir . '/excel-logo.webp', ['hyperlink' => 'https://github.com/aVadim483/fast-excel-writer']);

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);

        $sheet = $this->excelReader->sheet();
        //$this->assertEquals(5, $sheet->countImages());

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


    public function testExcelWriterSingleValue()
    {
        $testFileName = __DIR__ . '/test_single.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel1 = Excel::create();
        $sheet1 = $excel1->sheet();
        $sheet1->setValue('B5', 'test');

        $excel2 = Excel::create();
        $sheet2 = $excel2->sheet();
        $sheet2->setValue('B5:C7', 'test');

        $excel3 = Excel::create();
        $sheet3 = $excel3->sheet();

        $this->excelReader = $this->saveCheckRead($excel1, $testFileName);
        $this->cells = $this->excelReader->readCells();
        $this->assertEquals('test', $this->cells['B5']);
        unlink($testFileName);
        $this->cells = [];

        $this->excelReader = $this->saveCheckRead($excel2, $testFileName);
        $this->cells = $this->excelReader->readCells();
        $this->assertEquals('test', $this->cells['B5']);
        unlink($testFileName);
        $this->cells = [];

        $sheet3->setValue([2, 5], 'test');
        $sheet3->cell('C3')->writeCell('C3');

        $this->excelReader = $this->saveCheckRead($excel3, $testFileName);
        $this->cells = $this->excelReader->readCells();
        $this->assertEquals('test', $this->cells['B5']);
        $this->assertEquals('C3', $this->cells['C3']);
        unlink($testFileName);
    }


    public function testCellAddresses()
    {
        $testFileName = __DIR__ . '/test_addresses.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->writeTo('RC1', 'test1');
        $excel->setR1C1(false);
        $sheet->writeTo('RC1', 'test2');

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->cells = $this->excelReader->readCells();
        $this->assertEquals('test1', $this->cells['B1']);
        $this->assertEquals('test2', $this->cells['RC1']);
        unlink($testFileName);
        $this->cells = [];
    }


    public function testOverStyles()
    {
        $testFileName = __DIR__ . '/test_styles1.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $sheet->setColDataStyle('C', ['format' => '#,##0']);
        $sheet->writeCell(123.456);
        $sheet->writeCell(123.456, ['format' => '#,##0.00']);
        $sheet->writeCell(123.456);
        $sheet->writeRow([1, 2, 3], ['format' => '#,##0.0'], [null, ['format' => '#,##0.000']]);

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $style = $this->getStyle('A1', true);
        $this->assertEquals('GENERAL', $style['format-pattern']);
        $style = $this->getStyle('B1', true);
        $this->assertEquals('#,##0.00', $style['format-pattern']);
        $style = $this->getStyle('C1', true);
        $this->assertEquals('#,##0', $style['format-pattern']);

        $style = $this->getStyle('A2', true);
        $this->assertEquals('#,##0.0', $style['format-pattern']);
        $style = $this->getStyle('B2', true);
        $this->assertEquals('#,##0.000', $style['format-pattern']);
        $style = $this->getStyle('C2', true);
        $this->assertEquals('#,##0.0', $style['format-pattern']);

        unlink($testFileName);
        $this->cells = [];
    }


    public function testStyleClass()
    {
        $testFileName = __DIR__ . '/test_styles2.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $styles = [
            'A1' => [
                'set' => (new Style())->setBorder(Style::BORDER_THIN),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => "thin",
                    'border-right-style' => "thin",
                    'border-top-style' => "thin",
                    'border-bottom-style' => "thin",
                    'border-diagonal-style' => NULL,
                    'border-left-color' => "#000000",
                    'border-right-color' => "#000000",
                    'border-top-color' => "#000000",
                    'border-bottom-color' => "#000000",
                ],
            ],
            'A2' => [
                'set' => (new Style())->setBorderLeft(Style::BORDER_THICK, '#f00'),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => "thick",
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                    'border-left-color' => "#FF0000",
                ],],
            'A3' => [
                'set' => (new Style())->setBorderRight(Style::BORDER_THICK),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => "thick",
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                    'border-right-color' => "#000000",
                ],],
            'A4' => [
                'set' => (new Style())->setBorderTop(Style::BORDER_THICK, '#0f0'),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => "thick",
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                    'border-top-color' => "#00FF00",
                ],],
            'A5' => [
                'set' => (new Style())->setBorderBottom(Style::BORDER_THICK, '#009'),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => "thick",
                    'border-diagonal-style' => NULL,
                    'border-bottom-color' => "#000099",
                ],],
            'A6' => [
                'set' => (new Style())->setFont('Times New Roman', 16, Style::FONT_STYLE_ITALIC, '#900'),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => "Times New Roman",
                    'font-charset' => '1',
                    'font-family' => '1',
                    'font-size' => "16",
                    'font-color' => "#990000",
                    'font-style-italic' =>    1,
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A7' => [
                'set' => (new Style())->setFontName('Times New Roman'),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => "Times New Roman",
                    'font-charset' => '1',
                    'font-family' => '1',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A8' => [
                'set' => (new Style())->setFontStyle(Style::FONT_STYLE_BOLD),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'font-style-bold' =>    1,
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A9' => [
                'set' => (new Style())->setFontStyleBold(),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'font-style-bold' =>    1,
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A10' => [
                'set' => (new Style())->setFontStyleItalic(),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => "GENERAL",
                    'format-category' => "",
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'font-style-italic' => 1,
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A11' => [
                'set' => (new Style())->setFontStyleUnderline(),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'font-style-underline' => 1,
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A12' => [
                'set' => (new Style())->setFontStyleStrikethrough(),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'font-style-strike' => 1,
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A13' => [
                'set' => (new Style())->setFontColor('#099'),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'font-color' => '#009999',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A14' => [
                'set' => (new Style())->setFillColor('#fff000'),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'solid',
                    'fill-color' => '#FFF000',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A15' => [
                'set' => (new Style())->setFillGradient('#fff000', '#fff'),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A16' => [
                'set' => (new Style())->setTextAlign(Style::TEXT_ALIGN_CENTER),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'format-align-horizontal' => 'center',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A17' => [
                'set' => (new Style())->setTextCenter(),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'format-align-horizontal' => 'center',
                    'format-align-vertical' => 'center',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A18' => [
                'set' => (new Style())->setAlignLeft(),
                'get' => [
                    'format-num-id' => 164,
                    'format-pattern' => 'GENERAL',
                    'format-category' => '',
                    'format-align-horizontal' => 'left',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
            'A19' => [
                'set' => (new Style())->setFormat('@'),
                'get' => [
                    'format-num-id' => 165,
                    'format-pattern' => "@",
                    'format-category' => '',
                    'font-name' => 'Calibri',
                    'font-charset' => '1',
                    'font-family' => '0',
                    'font-size' => '11',
                    'fill-pattern' => 'none',
                    'border-left-style' => NULL,
                    'border-right-style' => NULL,
                    'border-top-style' => NULL,
                    'border-bottom-style' => NULL,
                    'border-diagonal-style' => NULL,
                ],],
        ];

        foreach ($styles as $cell => $style) {
            $sheet
                ->writeTo($cell, 12345, $style['set']->toArray())
                ->nextRow();
        }

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        foreach ($styles as $cell => $style) {
            $getStyles = $this->getStyle($cell, true);
            $this->assertEquals($getStyles, $style['get']);
        }

        unlink($testFileName);
        $this->cells = [];
    }

    public function testCharts()
    {
        $testFileName = __DIR__ . '/test_charts.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $data = [
            ['',	2010,	2011,	2012],
            ['Q1',   12,   15,		21],
            ['Q2',   56,   73,		86],
            ['Q3',   52,   61,		69],
            ['Q4',   30,   32,		0],
        ];

        foreach ($data as $row) {
            $sheet->writeRow($row);
        }

        $chartTypes1 = [
            Chart::TYPE_BAR              ,
            Chart::TYPE_BAR_STACKED      ,
            Chart::TYPE_COLUMN           ,
            Chart::TYPE_COLUMN_STACKED   ,
            Chart::TYPE_LINE             ,
            Chart::TYPE_LINE_STACKED     ,
            Chart::TYPE_LINE_3D          ,
            Chart::TYPE_LINE_3D_STACKED  ,
            Chart::TYPE_AREA             ,
            Chart::TYPE_AREA_STACKED     ,
            Chart::TYPE_AREA_3D          ,
            Chart::TYPE_AREA_3D_STACKED  ,
        ];
        $chartTypes2 = [
            Chart::TYPE_PIE              ,
            Chart::TYPE_PIE_3D           ,
            Chart::TYPE_DONUT            ,
        ];

        $dataSeries = [
            // key - cell with name of data series
            // value - range with data series
            'B1' => 'B2:B5',
            'C1' => 'c2:c5',
            'D1' => 'd2:d5',
        ];
        foreach ($chartTypes1 as $charType) {
            $chart = Chart::make($charType, $charType, $dataSeries)
                // X axis tick values
                ->setCategoryAxisLabels('A2:A5')
                // Position of legend
                ->setLegendPosition(Legend::POSITION_TOPRIGHT)
            ;
            $sheet->addChart('A7:H20', $chart);
        }

        foreach ($chartTypes2 as $charType) {
            $chart = Chart::make($charType, $charType, ['b6:d6'])
                // X axis tick values
                ->setCategoryAxisLabels('A2:A5')
                // Position of legend
                ->setLegendPosition(Legend::POSITION_TOPRIGHT)
            ;
            $sheet->addChart('A7:H20', $chart);
        }

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->cells = $this->excelReader->readCells();
        $this->assertEquals(2010, $this->cells['B1']);
        unlink($testFileName);
        $this->cells = [];
    }

    public function testConditional()
    {
        $testFileName = __DIR__ . '/test_conditional.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $value = 10;
        $style = [Style::FONT_COLOR => '#900', Style::FILL_COLOR => '#f99'];

        $conditional = [];

        $conditional[] = Conditional::make('=', $value, $style);
        $conditional[] = Conditional::equals($value, $style); // the same result

        $conditional[] = Conditional::make('!=', $value, $style);
        $conditional[] = Conditional::make('<>', $value, $style);
        $conditional[] = Conditional::notEquals($value, $style);

        $conditional[] = Conditional::make('>', $value, $style);
        $conditional[] = Conditional::greaterThan($value, $style);

        $conditional[] = Conditional::make('>=', $value, $style);
        $conditional[] = Conditional::greaterThanOrEqual($value, $style);

        $conditional[] = Conditional::make('<', $value, $style);
        $conditional[] = Conditional::lessThan($value, $style);

        $conditional[] = Conditional::make('<=', $value, $style);
        $conditional[] = Conditional::lessThanOrEqual($value, $style);

        $conditional[] = Conditional::make('between', [10, 50], $style);
        $conditional[] = Conditional::between([10, 50], $style);

        $conditional[] = Conditional::make('!between', [10, 50], $style);
        $conditional[] = Conditional::notBetween([10, 50], $style);

        $conditional[] = Conditional::make('=', 'Hello', $style);
        $conditional[] = Conditional::contains('Hello', $style);
        $conditional[] = Conditional::notContains('Hello', $style);
        $conditional[] = Conditional::beginsWith('Hello', $style);
        $conditional[] = Conditional::endsWith('Hello', $style);

        $conditional[] = Conditional::expression('=B10+SUM(C3:D8)', $style);
        $conditional[] = Conditional::isEmpty('B10', $style);

        $conditional[] = Conditional::colorScale('f00', '0f0');
        $conditional[] = Conditional::colorScale('f00', '0f0', '00f');
        $conditional[] = Conditional::colorScaleMax('f00');
        $conditional[] = Conditional::colorScaleMin('f00');
        $conditional[] = Conditional::colorScaleNum([10, 20], 'f00', '0f0', '00f');

        $conditional[] = Conditional::dataBar('f00')
            ->setGradient(false)
            ->setShowValue(false)
            ->setDirectionRtl(true);

        $conditional[] = Conditional::aboveAverage($style);
        $conditional[] = Conditional::belowAverage($style);

        $conditional[] = Conditional::uniqueValues($style);
        $conditional[] = Conditional::duplicateValues($style);

        $conditional[] = Conditional::top(5, $style);
        $conditional[] = Conditional::topPercent(5, $style);

        $sheet->addConditionalFormatting('a1:a' . Excel::MAX_ROW, $conditional);

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        unlink($testFileName);
        $this->cells = [];
    }


    public function testMergedCells()
    {
        $testFileName = __DIR__ . '/test_merge.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $sheet->mergeCells('A1:C2');
        $sheet->mergeCells('B3:E3');
        $merged = $sheet->getMergedCells();
        $this->assertEquals('A1:C2', $merged[0]);
        $this->assertEquals('B3:E3', $merged[1]);

        $sheet->mergeCells('D3:F3', Sheet::MERGE_REPLACE);
        $merged = $sheet->getMergedCells();
        $this->assertEquals('D3:F3', $merged[1]);

        $sheet->mergeCells('E3:G4', Sheet::MERGE_KEEP);
        $merged = $sheet->getMergedCells();
        $this->assertEquals('D3:F3', $merged[1]);

        $sheet->mergeCells('E3:G4', Sheet::MERGE_NO_CHECK);
        $merged = $sheet->getMergedCells();
        $this->assertEquals('A1:C2', $merged[0]);
        $this->assertEquals('D3:F3', $merged[1]);
        $this->assertEquals('E3:G4', $merged[2]);

        $excel->save($testFileName);
        unlink($testFileName);
        $this->cells = [];

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $sheet->writeTo('a1:c3', 'aaa');
        $sheet->writeTo('b4:e4', 'bbb');
        $merged = $sheet->getMergedCells();
        $this->assertEquals('A1:C3', $merged[0]);
        $this->assertEquals('B4:E4', $merged[1]);

        $sheet->writeTo('C4:D5', 'ccc', null, Sheet::MERGE_REPLACE);
        $merged = $sheet->getMergedCells();
        $this->assertEquals('A1:C3', $merged[0]);
        $this->assertEquals('C4:D5', $merged[1]);

        $sheet->writeTo('C5:D5', 'ddd', null, Sheet::MERGE_KEEP);
        $merged = $sheet->getMergedCells();
        $this->assertEquals('A1:C3', $merged[0]);
        $this->assertEquals('C4:D5', $merged[1]);

        $sheet->writeTo('C5:D5', 'eee', null, Sheet::MERGE_NO_CHECK);
        $merged = $sheet->getMergedCells();
        $this->assertEquals('A1:C3', $merged[0]);
        $this->assertEquals('C4:D5', $merged[1]);
        $this->assertEquals('C5:D5', $merged[2]);

        $excel->save($testFileName);
        unlink($testFileName);
        $this->cells = [];
    }

    public function testMergeException1()
    {
        $testFileName = __DIR__ . '/test_merge1.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $this->expectException(\avadim\FastExcelWriter\Exceptions\ExceptionAddress::class);
        $sheet->mergeCells('A1:C3');
        $sheet->mergeCells('B2:D12');

        unlink($testFileName);
        $this->cells = [];
    }

    public function testMergeException2()
    {
        $testFileName = __DIR__ . '/test_merge2.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $this->expectException(\avadim\FastExcelWriter\Exceptions\ExceptionAddress::class);
        $sheet->writeTo('A1:C3', 'aaa');
        $sheet->writeTo('B2:D12', 'bbb');

        unlink($testFileName);
        $this->cells = [];
    }

    protected function rmdir($tempDir)
    {
        if (is_dir($tempDir)) {
            foreach (glob($tempDir . '/*.tmp') as $file) {
                unlink($file);
            }
            clearstatcache();
            rmdir($tempDir);
        }
    }
/*
    public function testStyles()
    {
        $testFileName = __DIR__ . '/styles.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create();
        $sheet = $excel->sheet();

        $style = (new Style())->setColor('red');
        $sheet->setColStyle('a', $style);
        $sheet->setColStyle(2, $style);
        $sheet->setColStyle('C:D', $style);
        $sheet->setColStyle(['E', 'F'], $style);

        $this->excelReader = $this->saveCheckRead($excel, $testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);

        $style = $this->getStyle('a7', true);
        var_dump($style);

        unlink($testFileName);

        $this->excelReader = null;
        $this->cells = [];

    }
*/
}
