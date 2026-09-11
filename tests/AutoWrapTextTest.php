<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Options;
use avadim\FastExcelWriter\RichText\RichText;
use avadim\FastExcelReader\Excel as ExcelReader;
use PHPUnit\Framework\TestCase;

/**
 * The wrap text is enabled automatically for multi-line strings
 */
final class AutoWrapTextTest extends TestCase
{
    protected array $savedFiles = [];

    protected function tearDown(): void
    {
        foreach ($this->savedFiles as $file) {
            if (file_exists($file)) {
                unlink($file);
            }
        }
        $this->savedFiles = [];
    }


    protected function save(Excel $excel, string $testFileName): void
    {
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }
        $this->savedFiles[] = $testFileName;
        $excel->save($testFileName);
        $this->assertTrue(ExcelReader::validate($testFileName, $errors));
    }


    protected function readXml(string $testFileName, string $entry): string
    {
        $zip = new ZipArchive();
        $zip->open($testFileName);
        $xml = $zip->getFromName($entry);
        $zip->close();

        return (string)$xml;
    }


    /**
     * Returns the <xf> element of the cell style
     *
     * @param string $testFileName
     * @param string $cellAddress
     *
     * @return string
     */
    protected function cellXf(string $testFileName, string $cellAddress): string
    {
        $sheetXml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');
        $this->assertEquals(1, preg_match('#<c r="' . $cellAddress . '"(?: s="(\d+)")?#', $sheetXml, $m), 'Cell ' . $cellAddress . ' not found');
        $styleIdx = empty($m[1]) ? 0 : (int)$m[1];

        $stylesXml = $this->readXml($testFileName, 'xl/styles.xml');
        $this->assertEquals(1, preg_match('#<cellXfs[^>]*>(.*?)</cellXfs>#s', $stylesXml, $m));
        preg_match_all('#<xf\b[^>]*?(?:/>|>.*?</xf>)#s', $m[1], $xfs);
        $this->assertArrayHasKey($styleIdx, $xfs[0]);

        return $xfs[0][$styleIdx];
    }


    /**
     * Returns widths of columns by their numbers
     *
     * @param string $testFileName
     *
     * @return array
     */
    protected function colWidths(string $testFileName): array
    {
        $sheetXml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');
        preg_match_all('#<col\b([^>]*)>#', $sheetXml, $cols);
        $widths = [];
        foreach ($cols[1] as $attributes) {
            if (preg_match('#\bmin="(\d+)"#', $attributes, $min) && preg_match('#\bwidth="([\d.]+)"#', $attributes, $width)) {
                $widths[(int)$min[1]] = (float)$width[1];
            }
        }

        return $widths;
    }


    public function testMultiLineStringIsWrapped()
    {
        $testFileName = __DIR__ . '/test_auto_wrap.xlsx';

        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();
        $sheet->writeRow(['single line', "two\nlines", 'single\nquoted', '="a"&CHAR(10)&"b"', "windows\r\nbreak"]);

        $this->save($excel, $testFileName);

        $this->assertStringNotContainsString('wrapText', $this->cellXf($testFileName, 'A1'));
        $this->assertStringContainsString('wrapText="true"', $this->cellXf($testFileName, 'B1'));
        // a backslash followed by "n" is not a line break
        $this->assertStringNotContainsString('wrapText', $this->cellXf($testFileName, 'C1'));
        $this->assertStringNotContainsString('wrapText', $this->cellXf($testFileName, 'D1'));
        $this->assertStringContainsString('wrapText="true"', $this->cellXf($testFileName, 'E1'));

        $cells = ExcelReader::open($testFileName)->readCells();
        $this->assertEquals("two\nlines", $cells['B1']);
        $this->assertEquals('single\nquoted', $cells['C1']);
    }


    /**
     * The case of the issue #141: a bold header with multi-line titles
     */
    public function testMultiLineHeaderWithStyle()
    {
        $testFileName = __DIR__ . '/test_auto_wrap_header.xlsx';

        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();
        $sheet->writeHeader(['ARTICULO', "TIENDA\n1"])->applyFontStyleBold();

        $this->save($excel, $testFileName);

        $xfA1 = $this->cellXf($testFileName, 'A1');
        $xfB1 = $this->cellXf($testFileName, 'B1');
        $this->assertStringNotContainsString('wrapText', $xfA1);
        $this->assertStringContainsString('wrapText="true"', $xfB1);

        // both cells keep the bold font
        $this->assertEquals(1, preg_match('#fontId="(\d+)"#', $xfA1, $fontA1));
        $this->assertEquals(1, preg_match('#fontId="(\d+)"#', $xfB1, $fontB1));
        $this->assertNotEquals('0', $fontA1[1]);
        $this->assertEquals($fontA1[1], $fontB1[1]);
    }


    public function testExplicitWrapTextWins()
    {
        $testFileName = __DIR__ . '/test_auto_wrap_explicit.xlsx';

        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();
        $sheet->writeRow(["a\nb"], ['text-wrap' => false]);
        $sheet->writeRow(["c\nd"])->applyTextWrap(false);
        $sheet->writeRow(['single line'], ['text-wrap' => true]);

        $this->save($excel, $testFileName);

        $this->assertStringNotContainsString('wrapText', $this->cellXf($testFileName, 'A1'));
        $this->assertStringNotContainsString('wrapText', $this->cellXf($testFileName, 'A2'));
        $this->assertStringContainsString('wrapText="true"', $this->cellXf($testFileName, 'A3'));
    }


    public function testAutoWrapTextCanBeDisabled()
    {
        $testFileName = __DIR__ . '/test_auto_wrap_disabled.xlsx';

        $books = [
            'option' => Excel::create(['Sheet1'], ['auto_wrap_text' => false]),
            'options class' => Excel::create(['Sheet1'], Options::create()->autoWrapText(false)),
            'setter' => Excel::create(['Sheet1'])->setAutoWrapText(false),
        ];
        foreach ($books as $name => $excel) {
            $this->assertFalse($excel->isAutoWrapText(), $name);
            $excel->sheet()->writeRow(["a\nb"]);
            $this->save($excel, $testFileName);
            $this->assertStringNotContainsString('wrapText', $this->cellXf($testFileName, 'A1'), $name);
        }
        $this->assertTrue(Excel::create()->isAutoWrapText());
    }


    public function testMultiLineRichTextIsWrapped()
    {
        $testFileName = __DIR__ . '/test_auto_wrap_rich.xlsx';

        $excel = Excel::create(['Sheet1']);
        $excel->sheet()->writeRow([new RichText("<b>bold</b>\nnormal"), new RichText('<b>bold</b> normal')]);

        $this->save($excel, $testFileName);

        $this->assertStringContainsString('wrapText="true"', $this->cellXf($testFileName, 'A1'));
        $this->assertStringNotContainsString('wrapText', $this->cellXf($testFileName, 'B1'));
    }


    /**
     * The auto width of a column with a wrapped multi-line text is calculated by its longest line
     */
    public function testAutoWidthByLongestLine()
    {
        $testFileName = __DIR__ . '/test_auto_wrap_width.xlsx';

        $excel = Excel::create(['Sheet1']);
        $sheet = $excel->sheet();
        $sheet->setColDataStyle('A:C', ['width' => 'auto']);
        $sheet->writeRow(["short\nmuch longer line", 'much longer line', "short\nmuch longer line"], [], [2 => ['text-wrap' => false]]);

        $this->save($excel, $testFileName);

        $widths = $this->colWidths($testFileName);
        $this->assertArrayHasKey(1, $widths);
        $this->assertArrayHasKey(2, $widths);
        $this->assertArrayHasKey(3, $widths);
        $this->assertEqualsWithDelta($widths[2], $widths[1], 0.001);
        // without the wrap text the whole text is in one line
        $this->assertGreaterThan($widths[2], $widths[3]);
    }
}
