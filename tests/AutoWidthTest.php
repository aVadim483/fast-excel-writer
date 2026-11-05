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
            [1234.5, '0.00', true, 'ru', '1234,50'],
            [1.995, '0.00', true, 'en', '2.00'],
            [-0.00001, '0.00', true, 'en', '-0.00'],
            [-0.0, '0.00', true, 'en', '0.00'],
            [0.125, '0.00%', true, 'en', '12.50%'],
            [-0.125, '0.00%', true, 'de', '-12,50%'],
            [-0.00001, '0.00%', true, 'en', '-0.00%'],
            [1e308, '0.00%', true, 'en', (string)1e308],
            [INF, '0.00', true, 'en', 'INF'],
            [-INF, '0.00%', true, 'en', '-INF'],
            [NAN, '0.00', true, 'en', 'NAN'],
            [12.5, '0.00', false, 'en', '12.5'],
            ['12.50', '0.00%', false, 'en', '12.50'],
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
            ['abc', '0;0;0;"hidden text"', false, 'en', 'hidden text'],
            ['abc', '0;0;0;', false, 'en', ''],
            [12, '# ?/?', true, 'en', '12'],
            [60, 'yyyy-mm-dd', true, 'en', '1900-02-29'],
            [61, 'yyyy-mm-dd', true, 'en', '1900-03-01'],
            [61, 'yyyymmdd', true, 'en', '19000301'],
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

    public function testSimpleNumericFormatsAgreeWithEquivalentColorSections(): void
    {
        $formatter = new DisplayFormatter();
        $values = [PHP_INT_MIN, PHP_INT_MAX, -1, 0, 1, -0.0, 0.005, -0.005, -0.00001, 1.995,
            5e-324, 1e20, 1e308, -1e308, INF, -INF, NAN, '12.50', false, true, null];
        for ($i = -100; $i <= 100; ++$i) {
            $values[] = $i / 317;
        }
        foreach (['en', 'de', 'ru', 'fr'] as $locale) {
            foreach (['0.00', '0.00%'] as $format) {
                foreach ($values as $value) {
                    foreach ([true, false] as $numeric) {
                        // Color metadata changes no display text and uses the general section formatter.
                        $this->assertSame($formatter->format($value, '[Black]' . $format, $numeric, $locale),
                            $formatter->format($value, $format, $numeric, $locale));
                    }
                }
            }
        }
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

    public function testMinimumWidthAndAutofilterLimit(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:B')->setColMinWidth('A', 20)->setAutoFilter('A1:B2');
        $sheet->writeRow(['i', str_repeat('W', 10000)]);
        $sheet->writeRow(['i', 'short']);
        $widths = $this->saveWidths($excel);
        $this->assertSame(20.0, $widths[1]);
        $this->assertSame(255.0, $widths[2]);
    }

    public function testSaturatedColumnsSkipMeasurementsAndResumeAfterWidthReset(): void
    {
        $calculator = new class extends Font {
            public int $calls = 0;

            public function calculateCellWidth($value, array $font, array $baseFont, bool $wrap): float
            {
                ++$this->calls;
                return parent::calculateCellWidth($value, $font, $baseFont, $wrap);
            }
        };
        $excel = Excel::create();
        $property = new ReflectionProperty($excel, 'autoWidthFont');
        $property->setAccessible(true);
        $property->setValue($excel, $calculator);
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:B')->setColMinWidth('B', 255)->setAutoFilter('A1:B3');
        $sheet->writeRow([str_repeat('W', 600), 'minimum width']);
        $sheet->nextRow();
        $this->assertSame(1, $calculator->calls);
        $sheet->writeRow(['later text', new RichText('Saved <b>rich</b> text')]);
        $sheet->nextRow();
        $this->assertSame(1, $calculator->calls);
        $sheet->setColWidth('A', 5);
        Font::setFontWidthFactor(0.8);
        $replacement = 'Automatic width after resetting';
        $sheet->writeRow([$replacement, false]);
        $widths = $this->saveWidths($excel);
        $this->assertSame(2, $calculator->calls);
        $base = $excel->styleManager->defaultFont['val'];
        $expected = (new Font())->calculateCellWidth($replacement, $base, $base, false) + 1;
        $this->assertEqualsWithDelta($expected, $widths[1], 1 / 256);
        $this->assertSame(255.0, $widths[2]);
        $cells = ExcelReader::open(end($this->files))->readCells();
        $this->assertSame('later text', $cells['A3']);
        $this->assertSame('Saved rich text', $cells['B3']);
        $this->assertSame($replacement, $cells['A5']);
    }

    public function testSaturatedColumnsStillInvokeNewCallbacksForEachValue(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A')->writeRow([str_repeat('W', 600)]);
        $sheet->nextRow();
        $seen = [];
        $excel->setAutoWidthFormatter(function ($value) use (&$seen) {
            $seen[] = $value;
            return null;
        });
        $sheet->writeRows([[12.5], [false], [['=1+1', 2]], ["short\ntext"]]);
        $widths = $this->saveWidths($excel);
        $this->assertSame([12.5, false, 2, "short\ntext"], $seen);
        $this->assertSame(255.0, $widths[1]);
        $cells = ExcelReader::open(end($this->files))->readCells();
        $this->assertEquals(12.5, $cells['A3']);
        $this->assertSame("short\ntext", $cells['A6']);
    }

    public function testSaturatedColumnsStillValidateCallbackResults(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A')->writeRow([str_repeat('W', 600)]);
        $sheet->nextRow();
        $excel->setAutoWidthFormatter(static function () { return 123; });
        $sheet->writeRow(['later text']);
        $this->expectException(UnexpectedValueException::class);
        $this->saveWidths($excel);
    }

    public function testFormatterCanReduceASaturatedColumnWidth(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A')->writeRow([str_repeat('W', 600)]);
        $sheet->nextRow();
        $label = 'Width from callback';
        $excel->setAutoWidthFormatter(function () use ($sheet, $label) {
            $sheet->setColWidth('A', 5);
            return $label;
        });
        $sheet->writeRow([false]);
        $widths = $this->saveWidths($excel);
        $base = $excel->styleManager->defaultFont['val'];
        $expected = (new Font())->calculateCellWidth($label, $base, $base, false);
        $this->assertEqualsWithDelta($expected, $widths[1], 1 / 256);
    }

    public function testWorkbookWidthContextIsReleased(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A')->writeRow(['text']);
        $this->saveWidths($excel);
        $reference = WeakReference::create($excel->getAutoWidthFont());
        unset($sheet, $excel);
        gc_collect_cycles();
        gc_collect_cycles(); // Cycles containing destructors can require a second collection.
        $this->assertTrue($reference->get() === null, 'The workbook must release its width calculator');
    }

    public function testWorkbookLocalesDoNotAffectEachOther(): void
    {
        $first = Excel::create([], ['locale' => 'en']);
        $second = Excel::create([], ['locale' => 'de']);
        $this->assertSame('1,234.50', $first->formatAutoWidthValue(1234.5, '#,##0.00', 'n_numeric', true));
        $this->assertSame('1.234,50', $second->formatAutoWidthValue(1234.5, '#,##0.00', 'n_numeric', true));
        $this->assertSame('1,234.50', $first->formatAutoWidthValue(1234.5, '#,##0.00', 'n_numeric', true));
    }

    public function testNumericDateWidthSamples(): void
    {
        $formatter = new DisplayFormatter();
        $this->assertSame('0000-00-00 00:00:00', $formatter->dateWidthSample('yyyy-mm-dd hh:mm:ss'));
        $this->assertSame('00/00/00', $formatter->dateWidthSample('m/d/yy'));
        $this->assertSame('date 00000000', $formatter->dateWidthSample('"date "yyyymmdd'));
        foreach (['mmmm d, yyyy', 'dddd', '[h]:mm', 'h:mm AM/PM', '[>=1]yyyy;mm', '0.00'] as $format) {
            $this->assertNull($formatter->dateWidthSample($format), $format);
            $this->assertNull($formatter->dateWidthSample($format), 'Cached ' . $format);
        }
    }

    public function testDatesUseTheMaskAndCallbacksStillReceiveTheValue(): void
    {
        $excel = Excel::create([], ['locale' => 'en']);
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:C');
        $sheet->writeRow(['2024-01-01 12:34:56', '2099-12-31 23:59:59', '0000-00-00 00:00:00'], [],
            [0 => ['format' => 'yyyy-mm-dd hh:mm:ss'], 1 => ['format' => 'yyyy-mm-dd hh:mm:ss']]);
        $widths = $this->saveWidths($excel);
        $this->assertEqualsWithDelta($widths[3], $widths[1], 1 / 256);
        $this->assertEqualsWithDelta($widths[3], $widths[2], 1 / 256);
        $seen = null;
        $excel->setAutoWidthFormatter(function ($value) use (&$seen) { $seen = $value; return 'custom'; });
        $this->assertSame('custom', $excel->formatAutoWidthValue(60, 'yyyy-mm-dd', 'n_date', true));
        $this->assertSame(60, $seen);
        $excel->setAutoWidthFormatter(null);
        $this->assertSame('1900-02-29', $excel->getAutoWidthFont()->formatDisplayValue(60, 'yyyy-mm-dd', true, 'en'));
        $this->assertSame('January', $excel->formatAutoWidthValue(1, 'mmmm', 'n_date', true));
    }

    public function testLongAndShortTextMeasurementsAgree(): void
    {
        $calculator = new class extends Font {
            public function pixels(string $text, array $font): float
            {
                return $this->measurePixels($text, $font);
            }
        };
        foreach (['Calibri', 'Arial', 'Tahoma', 'Times New Roman'] as $name) {
            foreach ([0, 1, 2, 3] as $face) {
                $font = ['name' => $name, 'size' => 13.5, 'style-bold' => (bool)($face & 1), 'style-italic' => (bool)($face & 2)];
                foreach (['WWW iii 0123 ', 'Текст Ёё 0123 ', '£ € ¥ ', "中文 e\xcc\x81 😀 ", "\t\n\0"] as $text) {
                    $this->assertEqualsWithDelta($calculator->pixels($text, $font) * 30,
                        $calculator->pixels(str_repeat($text, 30), $font), 0.000001, $name . ':' . $face);
                }
            }
        }
    }

    public function testMeasurementCacheCanReuseTabularDigitPatterns(): void
    {
        $digits = strpos(\avadim\FastExcelWriter\Style\FontMetrics::CHARACTERS, '0');
        foreach (\avadim\FastExcelWriter\Style\FontMetrics::WIDTHS as $faces) {
            foreach ($faces as $widths) {
                $this->assertCount(1, array_unique(array_slice($widths, $digits, 10)), 'Digit normalization requires tabular metrics');
            }
        }
        $font = new Font();
        $this->assertSame($font->calculateTextWidth('Calibri', 11, 'Row 12345'), $font->calculateTextWidth('Calibri', 11, 'Row 98765'));
        $this->assertGreaterThan($font->calculateTextWidth('Calibri', 11, 'Row 12345'), $font->calculateTextWidth('Calibri', 11, 'Row 123456'));
        $this->assertGreaterThan($font->calculateTextWidth('Calibri', 11, '12345'), $font->calculateTextWidth('Calibri', 11, '12345', '0%'));
    }

    public function testIntegerWidthsPreserveSignsFontsAndFactors(): void
    {
        $calculator = new Font();
        foreach ([['name' => 'Calibri', 'size' => 11], ['name' => 'Arial', 'size' => 16, 'style-bold' => true]] as $font) {
            foreach ([[], ['name' => 'Tahoma', 'size' => 14]] as $base) {
                foreach ([1.0, 1.5] as $factor) {
                    Font::setFontWidthFactor($factor);
                    foreach ([0, 1, -1, 9, 10, -10, 999999999999999, PHP_INT_MIN, PHP_INT_MAX] as $value) {
                        $this->assertSame($calculator->calculateCellWidth((string)$value, $font, $base, false),
                            $calculator->calculateIntegerWidth($value, $font, $base));
                    }
                }
            }
        }
    }

    public function testIntegerFastPathPreservesFormatsAndCallbacks(): void
    {
        foreach (['en', 'de', 'ru'] as $locale) {
            $excel = Excel::create([], ['locale' => $locale]);
            $sheet = $excel->sheet();
            $sheet->setColWidthAuto('A:H');
            $sheet->writeRow([12, '12', 12, '000012', 12, '12', 12, '1200%'], [],
                [2 => ['format' => '000000'], 4 => ['format' => '@'], 6 => ['format' => '0%']]);
            $widths = $this->saveWidths($excel);
            foreach ([[1, 2], [3, 4], [5, 6], [7, 8]] as $pair) {
                $this->assertSame($widths[$pair[0]], $widths[$pair[1]]);
            }
        }
        $seen = [];
        $excel = Excel::create();
        $excel->setAutoWidthFormatter(function ($value) use (&$seen) { $seen[] = $value; return 'custom integer'; });
        $excel->sheet()->setColWidthAuto('A:B')->writeRow([12, 0]);
        $this->saveWidths($excel);
        $this->assertSame([12, 0], $seen);
    }

    public function testBooleanWidthsPreserveFontAndFactorChanges(): void
    {
        $calculator = new Font();
        foreach ([['name' => 'Calibri', 'size' => 11], ['name' => 'Arial', 'size' => 16, 'style-italic' => true]] as $font) {
            foreach ([[], ['name' => 'Tahoma', 'size' => 14]] as $base) {
                foreach ([1.0, 1.5] as $factor) {
                    Font::setFontWidthFactor($factor);
                    foreach ([true, false, true] as $value) {
                        $this->assertSame($calculator->calculateCellWidth($value ? 'TRUE' : 'FALSE', $font, $base, false),
                            $calculator->calculateBooleanWidth($value, $font, $base));
                        $this->assertSame($calculator->calculateCellWidth('12', $font, $base, false),
                            $calculator->calculateIntegerWidth(12, $font, $base));
                    }
                }
            }
        }
    }

    public function testBooleansIgnoreNumberMasksAndPreserveCallbacks(): void
    {
        $excel = Excel::create();
        $values = [];
        $styles = [];
        foreach (['GENERAL', '@', '0.00', 'yyyy-mm-dd', '0;0;0;"custom"'] as $format) {
            foreach ([true, false] as $value) {
                $styles[count($values)] = ['format' => $format];
                $values[] = $value;
                $values[] = $value ? 'TRUE' : 'FALSE';
            }
        }
        $excel->sheet()->setColWidthAuto('A:T')->writeRow($values, [], $styles);
        $widths = $this->saveWidths($excel);
        for ($i = 1; $i <= 20; $i += 2) {
            $this->assertSame($widths[$i + 1], $widths[$i]);
        }
        $seen = [];
        $excel = Excel::create();
        $excel->setAutoWidthFormatter(function ($value) use (&$seen) {
            if (is_bool($value)) {
                $seen[] = $value;
                return $value ? 'custom true' : 'custom false';
            }
            return null;
        });
        $excel->sheet()->setColWidthAuto('A:B')->writeRows([[false, 'custom false'], [true, 'custom true']]);
        $widths = $this->saveWidths($excel);
        $this->assertSame([false, true], $seen);
        $this->assertSame($widths[1], $widths[2]);
    }

    public function testReadyBooleanWidthsPreserveInheritanceAndFilters(): void
    {
        $calculator = new Font();
        foreach ([['name' => 'Calibri', 'size' => 11], ['name' => 'Tahoma', 'size' => 18, 'style-bold' => true, 'style-italic' => true]] as $base) {
            foreach ([[], ['size' => 16], ['style-bold' => false, 'name' => 'Arial'], ['size' => 1000]] as $overrides) {
                foreach ([1.0, 1.5, 0.0, 1.0] as $factor) {
                    Font::setFontWidthFactor($factor);
                    foreach ([false, true] as $value) {
                        $expected = $calculator->calculateCellWidth($value ? 'TRUE' : 'FALSE', array_replace($base, $overrides), $base, false);
                        $this->assertSame($expected, $calculator->calculateBooleanStyleWidth($value, $overrides, $base));
                        $this->assertSame(min(255.0, $expected + 1), $calculator->calculateBooleanStyleWidth($value, $overrides, $base, true));
                        $direct = $calculator->calculateCellWidth($value ? 'TRUE' : 'FALSE', $overrides, $base, false);
                        $this->assertSame($direct, $calculator->calculateBooleanWidth($value, $overrides, $base));
                    }
                }
            }
        }
    }

    public function testBooleanReadyPathPreservesColumnLimitsAndZeroWidths(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:D')->setColMinWidth('A:B', 20)->setAutoFilter('A1:D2');
        $sheet->writeRows([[false, 'FALSE', true, 'TRUE'], [true, 'TRUE', false, 'FALSE']]);
        $widths = $this->saveWidths($excel);
        $this->assertSame(20.0, $widths[1]);
        $this->assertSame($widths[1], $widths[2]);
        $this->assertSame($widths[3], $widths[4]);
        Font::setFontWidthFactor(0.0);
        $excel = Excel::create();
        $excel->sheet()->setColWidthAuto('A:B')->writeRow([false, true]);
        $widths = $this->saveWidths($excel);
        $this->assertArrayHasKey(1, $widths);
        $this->assertArrayHasKey(2, $widths);
        $this->assertSame(0.0, $widths[1]);
        $this->assertSame(0.0, $widths[2]);
    }

    public function testBooleanCacheDoesNotBypassNewCallbacks(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A:B')->writeRow([false, 'FALSE']);
        $sheet->nextRow();
        $seen = [];
        $excel->setAutoWidthFormatter(function ($value) use (&$seen) {
            if (is_bool($value)) {
                $seen[] = $value;
                return 'A longer callback label';
            }
            return null;
        });
        $sheet->writeRow([true, 'A longer callback label']);
        $sheet->nextRow();
        $excel->setAutoWidthFormatter(null);
        $sheet->writeRow([false, 'FALSE']);
        $widths = $this->saveWidths($excel);
        $this->assertSame([true], $seen);
        $this->assertSame($widths[1], $widths[2]);
    }

    public function testReadyBooleanCacheMemoryIsBounded(): void
    {
        $calculator = new Font();
        $base = ['name' => 'Calibri', 'size' => 11];
        for ($i = 0; $i < 3000; ++$i) {
            $calculator->calculateBooleanStyleWidth(true, ['size' => 10 + $i / 1000], $base);
        }
        $memory = memory_get_usage();
        for ($i = 3000; $i < 15000; ++$i) {
            $calculator->calculateBooleanStyleWidth(false, ['size' => 10 + $i / 1000], $base, true);
        }
        $this->assertLessThan(1048576, memory_get_usage() - $memory);
    }
}
