<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\Font;
use avadim\FastExcelReader\Excel as ExcelReader;
use PHPUnit\Framework\TestCase;

final class AutoWidthMultilineTest extends TestCase
{
    protected function tearDown(): void { Font::setFontWidthFactor(1.0); }

    public function testPlainWrappedTextUsesTheLongestLineWithOnePadding(): void
    {
        $calculator = new Font();
        $values = [["i\nWWWW", 'WWWW'], ["WWWW\r\ni", 'WWWW'],
            ["short\nДлинная строка текста\nend", 'Длинная строка текста'],
            ["\n\nWWWW\n", 'WWWW'], ["中文\n中文中文", '中文中文'], ["i\n🙂🙂🙂🙂", '🙂🙂🙂🙂'],
            ["é\néééé", 'éééé'], ["a\rb\nWWWW", 'WWWW']];
        foreach ([1.0, 0.8, 1.5, 0.0] as $factor) {
            Font::setFontWidthFactor($factor);
            foreach ([['name' => 'Calibri', 'size' => 11], ['name' => 'Arial', 'size' => 16, 'style-bold' => true]] as $font) {
                foreach ([[], ['name' => 'Tahoma', 'size' => 14]] as $base) {
                    foreach ($values as [$text, $longest]) {
                        $expected = (new Font())->calculateCellWidth($longest, $font, $base, false);
                        // PHP 7.4's existing empty-line metric path may emit an mb_ord warning.
                        $this->assertSame($expected, @$calculator->calculateCellWidth($text, $font, $base, true));
                        $this->assertSame($expected, @$calculator->calculateCellWidth($text, $font, $base, true));
                    }
                }
            }
        }
    }

    public function testNonStringValuesKeepTheirExistingStringConversionPath(): void
    {
        $value = new class {
            public function __toString(): string { return "i\nWWWW"; }
        };
        $calculator = new Font();
        $this->assertSame($calculator->calculateCellWidth('WWWW', [], [], false),
            $calculator->calculateCellWidth($value, [], [], true));
    }

    public function testExportPreservesMultilineValuesAndCallsTheFormatterPerCell(): void
    {
        $excel = Excel::create([], ['shared_string' => false]);
        $sheet = $excel->sheet()->setColWidthAuto('A');
        $seen = [];
        $excel->setAutoWidthFormatter(function ($value) use (&$seen) { $seen[] = $value; return null; });
        $values = ["Short\nLonger text", "Short\r\nLonger text", "Русский текст\n中文🙂"];
        foreach ($values as $value) { $sheet->writeRow([$value]); }
        $path = tempnam(sys_get_temp_dir(), 'few-multiline-');
        try {
            $excel->save($path);
            $this->assertTrue(ExcelReader::validate($path, $errors));
            $this->assertSame($values, $seen);
            $zip = new ZipArchive(); $zip->open($path);
            $xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
            $zip->close();
            foreach ($values as $index => $value) {
                $this->assertSame(str_replace("\r", '_x000D_', $value), (string)$xml->sheetData->row[$index]->c->is->t);
            }
        }
        finally { unlink($path); }
    }
}
