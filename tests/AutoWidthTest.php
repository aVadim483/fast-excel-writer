<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\Font;
use avadim\FastExcelWriter\Style\WidthCache;
use avadim\FastExcelWriter\Style\DisplayFormatter;
use avadim\FastExcelWriter\RichText\RichText;
use avadim\FastExcelReader\Excel as ExcelReader;
use PHPUnit\Framework\TestCase;

final class AutoWidthTest extends TestCase
{
    protected array $files = [];

    protected function tearDown(): void
    {
        Font::setFontWidthFactor(1.0);
        foreach ($this->files as $file) {
            if (is_file($file)) {
                unlink($file);
            }
        }
    }

    protected function saveWidths(Excel $excel): array
    {
        $file = tempnam(sys_get_temp_dir(), 'few-width-');
        $this->files[] = $file;
        $excel->save($file);
        $this->assertTrue(ExcelReader::validate($file, $errors));
        $zip = new ZipArchive();
        $zip->open($file);
        $xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
        $zip->close();
        $widths = [];
        foreach ($xml->cols->col as $col) {
            for ($index = (int)$col['min']; $index <= (int)$col['max']; ++$index) {
                $widths[$index] = (float)$col['width'];
            }
        }
        return $widths;
    }

    public function testRichTextAndPlainTextHaveTheSameWidth(): void
    {
        foreach ([false, true] as $shared) {
            $excel = Excel::create([], ['shared_string' => $shared]);
            $sheet = $excel->sheet();
            $sheet->setColWidthAuto('A:D');
            $rich = new RichText();
            $rich->addText('Sample ')->addText('text & <literal>');
            $sheet->writeRow([$rich, 'Sample text & <literal>', new RichText(''), '0']);
            $widths = $this->saveWidths($excel);
            $this->assertEqualsWithDelta($widths[2], $widths[1], 0.001);
            $this->assertGreaterThan(0, $widths[4]);
            $cells = ExcelReader::open(end($this->files))->readCells();
            $this->assertSame('Sample text & <literal>', $cells['A1']);
        }
    }

    public function testRichTextLinesAndFragmentSizes(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:D');
        $rich = new RichText();
        $rich->addText("short\r\nlonger ")->addText('line');
        $large = new RichText();
        $large->addText('longer line')->setSize(22);
        $sheet->writeRow([$rich, 'longer line', $large, new RichText('longer line')]);
        $widths = $this->saveWidths($excel);
        $this->assertEqualsWithDelta($widths[2], $widths[1], 0.001);
        $this->assertGreaterThan($widths[4], $widths[3]);
    }

    public function testZerosBooleansAndFormulaResults(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:F');
        $sheet->writeRow([0, '0', false, 'FALSE', ['=1-1', 0], ['=1+1', null]]);
        $widths = $this->saveWidths($excel);
        $this->assertGreaterThan(0, $widths[1]);
        $this->assertEqualsWithDelta($widths[1], $widths[2], 0.001);
        $this->assertEqualsWithDelta($widths[1], $widths[5], 0.001);
        $this->assertEqualsWithDelta($widths[3], $widths[4], 0.001);
        $this->assertArrayNotHasKey(6, $widths);
    }

    public function testFactorChangesAffectCachedWidths(): void
    {
        $width = Font::calcTextWidth('Calibri', 11, 'cached text');
        Font::setFontWidthFactor(1.5);
        $this->assertEqualsWithDelta($width * 1.5, Font::calcTextWidth('Calibri', 11, 'cached text'), 1 / 256);
    }

    public function testCacheEvictionAndKeyBudgets(): void
    {
        $cache = new WidthCache(2, 6, 4);
        $cache->put('aa', 0);
        $cache->put('bb', 2);
        $this->assertSame(0, $cache->get('aa'));
        $cache->put('cccc', 3);
        $this->assertNull($cache->get('aa'));
        $this->assertSame(2, $cache->get('bb'));
        $cache->put('long-key', 4);
        $this->assertNull($cache->get('long-key'));
        for ($i = 0; $i < 20; ++$i) {
            $cache->put((string)$i, $i);
        }
        $this->assertSame(19, $cache->get('19'));
        $this->assertNull($cache->get('0'));
    }

    public function testUniqueValuesDoNotKeepGrowingMemory(): void
    {
        $font = new Font();
        for ($i = 0; $i < 10000; ++$i) {
            $font->calculateTextWidth('Calibri', 11, (string)$i);
        }
        $memory = memory_get_usage();
        for ($i = 10000; $i < 60000; ++$i) {
            $font->calculateTextWidth('Calibri', 11, (string)$i);
        }
        $this->assertLessThan(1048576, memory_get_usage() - $memory);
    }

    public function testWorkbookCachesAreIndependent(): void
    {
        $first = Excel::create();
        $second = Excel::create();
        $this->assertNotSame($first->getAutoWidthFont(), $second->getAutoWidthFont());
        $this->assertSame($first->getAutoWidthFont(), $first->getAutoWidthFont());
    }

    public function displayValues(): array
    {
        return [
            [0.25, '0%', true, 'en', '25%'],
            [12, '000000', true, 'en', '000012'],
            [1234.5, '0.00', true, 'en', '1234.50'],
            [0.000001234, '0.00', true, 'en', '0.00'],
            ['000012', '@', false, 'en', '000012'],
            ['12', '000000', false, 'en', '12'],
            [1234.5, '#,##0.00', true, 'en', '1,234.50'],
            [1234.5, '#,##0.00', true, 'de', '1.234,50'],
            [1234.5, '#,##0.00', true, 'ru', "1\xc2\xa0234,50"],
            [1234567, '0.0,,', true, 'en', '1.2'],
            [1234567, '0.00E+00', true, 'en', '1.23E+06'],
            [0.0000123, '0.0E+00', true, 'en', '1.2E-05'],
            [12.3, '0.##', true, 'en', '12.3'],
            [-12.5, '0.00;(0.00);"zero"', true, 'en', '(12.50)'],
            [0, '0.00;(0.00);"zero"', true, 'en', 'zero'],
            [0, '0.00;0.00;', true, 'en', ''],
            [5, '[>=10]0.0;[<0](0);"small"', true, 'en', 'small'],
            [10, '[>=10]0.0;[<0](0);"small"', true, 'en', '10.0'],
            [12, '"units; "0', true, 'en', 'units; 12'],
            [12, '\\$0', true, 'en', '$12'],
            [12, '[$€-407]0.00', true, 'en', '€12,00'],
            ['abc', '0;0;0;"text: "@', false, 'en', 'text: abc'],
            [12, '# ?/?', true, 'en', '12'],
            [60, 'yyyy-mm-dd', true, 'en', '1900-02-29'],
            [61, 'yyyy-mm-dd', true, 'en', '1900-03-01'],
            [0, 'yyyy-mm-dd', true, 'en', '1900-01-00'],
            [1.5, '[h]:mm', true, 'en', '36:00'],
            [0.75, 'h:mm AM/PM', true, 'en', '6:00 PM'],
            [61.520833333, 'yyyy-mm-dd hh:mm', true, 'en', '1900-03-01 12:30'],
            [true, '0', false, 'en', 'TRUE'],
        ];
    }

    /** @dataProvider displayValues */
    public function testDisplayFormatting($value, string $format, bool $numeric, string $locale, string $expected): void
    {
        $formatter = new DisplayFormatter();
        $this->assertSame($expected, $formatter->format($value, $format, $numeric, $locale));
    }

    public function testWidthsUseDisplayedTextOnly(): void
    {
        $excel = Excel::create([], ['locale' => 'en']);
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:D');
        $sheet->writeRow([0.000001234, '0.00', 12, '000012'], [], [0 => ['format' => '0.00'], 2 => ['format' => '000000']]);
        $widths = $this->saveWidths($excel);
        $this->assertEqualsWithDelta($widths[2], $widths[1], 0.001);
        $this->assertEqualsWithDelta($widths[4], $widths[3], 0.001);
    }

    public function testFormatterCallbackIsNotCachedAndDoesNotChangeCells(): void
    {
        $calls = 0;
        $excel = Excel::create();
        $excel->setAutoWidthFormatter(function ($value, $format, $type, $locale) use (&$calls) {
            ++$calls;
            return $value === 1 ? 'custom text' : null;
        });
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:B');
        $sheet->writeRow([1, 'custom text']);
        $sheet->writeRow([1, 'custom text']);
        $widths = $this->saveWidths($excel);
        $this->assertSame(4, $calls);
        $this->assertEqualsWithDelta($widths[1], $widths[2], 0.001);
        $cells = ExcelReader::open(end($this->files))->readCells();
        $this->assertEquals(1, $cells['A1']);
    }

    public function testExcelAutoFitReferences(): void
    {
        $fixture = json_decode(ltrim(file_get_contents(__DIR__ . '/auto-width-excel.json'), "\xef\xbb\xbf"), true);
        $calculator = new Font();
        foreach ($fixture['cases'] as $case) {
            $text = $calculator->formatDisplayValue($case['value'], $case['format'], is_numeric($case['value']), $fixture['locale']);
            if (isset($case['display'])) {
                $this->assertSame($case['display'], $text);
            }
            $font = ['name' => $case['font'], 'size' => $case['size'], 'style-bold' => $case['bold'], 'style-italic' => $case['italic']];
            $width = $calculator->calculateCellWidth($text, $font, ['name' => $fixture['base_font'], 'size' => $fixture['base_size']], false);
            $this->assertEqualsWithDelta($case['xml_width'], $width, max(1.15, $case['xml_width'] * 0.05), json_encode($case, JSON_UNESCAPED_UNICODE));
        }
    }

    public function testWorkbookFontAndFragmentOverrides(): void
    {
        $calculator = new Font();
        $normal = ['name' => 'Arial', 'size' => 11];
        $largeBase = ['name' => 'Arial', 'size' => 22];
        $text = 'Sample WWW text';
        $width = $calculator->calculateCellWidth($text, $normal, $normal, false);
        $this->assertLessThan($width, $calculator->calculateCellWidth($text, $normal, $largeBase, false));
        $rich = new RichText();
        $rich->addText($text)->setBold(false)->setItalic(false);
        $this->assertSame($width, $calculator->calculateCellWidth($rich, $normal + ['style-bold' => true, 'style-italic' => true], $normal, false));
        $this->assertGreaterThan($calculator->calculateCellWidth('iiii', $normal, $normal, false), $calculator->calculateCellWidth('iiii', $normal + ['style-bold' => true], $normal, false));
        $this->assertSame($calculator->calculateTextWidth('Calibri', 11, '中文'), $calculator->calculateTextWidth('Unknown font', 11, '中文'));
        $this->assertLessThanOrEqual(255, $calculator->calculateTextWidth('Calibri', 11, str_repeat('W', 10000)));
    }
}
