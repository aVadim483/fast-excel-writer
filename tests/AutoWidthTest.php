<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\Font;
use avadim\FastExcelWriter\Style\WidthCache;
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
        $this->assertEqualsWithDelta($width * 1.5, Font::calcTextWidth('Calibri', 11, 'cached text'), 0.000001);
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
}
