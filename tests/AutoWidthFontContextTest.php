<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\Font;
use PHPUnit\Framework\TestCase;

final class AutoWidthFontContextTest extends TestCase
{
    protected function tearDown(): void
    {
        Font::setFontWidthFactor(1.0);
    }

    public function testInheritanceTracksChangesAndReturnedArraysAreIndependent(): void
    {
        $calculator = new Font();
        $base = ['name' => 'Arial', 'size' => 11, 'style-bold' => true];
        $overrides = ['size' => 16, 'style-bold' => false];
        for ($i = 0; $i < 3; ++$i) {
            $this->assertSame(array_replace($base, $overrides), $calculator->resolveCellFont($base, $overrides));
        }
        $resolved = $calculator->resolveCellFont($base, $overrides);
        $resolved['name'] = 'Tahoma';
        $this->assertSame(array_replace($base, $overrides), $calculator->resolveCellFont($base, $overrides));
        $base['name'] = 'Times New Roman';
        unset($overrides['size']);
        $this->assertSame(array_replace($base, $overrides), $calculator->resolveCellFont($base, $overrides));
        $this->assertSame([], $calculator->resolveCellFont([], []));
    }

    public function testReferencedFontValuesInvalidateMeasurementContexts(): void
    {
        $calculator = new Font();
        $size = 11;
        $base = ['name' => 'Calibri', 'size' => &$size];
        for ($i = 0; $i < 2; ++$i) {
            $font = $calculator->resolveCellFont($base, []);
            $calculator->calculateCellWidth('WWWW iii', $font, [], false);
        }
        $size = 22;
        $font = $calculator->resolveCellFont($base, []);
        $this->assertSame(array_replace($base, []), $font);
        $expected = (new Font())->calculateCellWidth('WWWW iii', ['name' => 'Calibri', 'size' => 22], [], false);
        $this->assertSame($expected, $calculator->calculateCellWidth('WWWW iii', $font, [], false));
    }

    public function testCompleteFontsStillInheritNewBaseProperties(): void
    {
        $calculator = new Font();
        $base = ['name' => 'Calibri', 'size' => 11, 'style-bold' => true];
        $font = ['size' => 16, 'name' => 'Arial', 'style-bold' => false, 'color' => 'FF0000'];
        $this->assertEquals(array_replace($base, $font), $calculator->resolveCellFont($base, $font));
        $base['style-italic'] = true;
        $this->assertEquals(array_replace($base, $font), $calculator->resolveCellFont($base, $font));
    }

    public function testAlternatingFontContextsAgreeWithFreshCalculators(): void
    {
        $calculator = new Font();
        foreach ([1.0, 1.4, 0.0, 1.0] as $factor) {
            Font::setFontWidthFactor($factor);
            foreach (['Calibri', 'Arial', 'Tahoma', 'Times New Roman', 'Unknown font'] as $name) {
                foreach ([11, 16] as $size) {
                    foreach ([0, 1, 2, 3] as $face) {
                        $font = ['name' => $name, 'size' => $size, 'style-bold' => (bool)($face & 1), 'style-italic' => (bool)($face & 2)];
                        foreach ([[], ['name' => 'Arial', 'size' => 13]] as $base) {
                            $fresh = new Font();
                            foreach (['Text 123 WWW iii', 'Текст Ёё 456', "line\nlonger line", '中文 é 😀'] as $text) {
                                $expected = $fresh->calculateCellWidth($text, $font, $base, true);
                                $this->assertSame($expected, $calculator->calculateCellWidth($text, $font, $base, true));
                                $this->assertSame($expected, $calculator->calculateCellWidth($text, $font, $base, true));
                            }
                        }
                    }
                }
            }
        }
    }

    public function testPreparedContextsDoNotRetainEveryFont(): void
    {
        $calculator = new Font();
        $base = ['name' => 'Calibri', 'size' => 11];
        for ($i = 0; $i < 6000; ++$i) {
            $font = $calculator->resolveCellFont($base, ['name' => 'Unknown ' . $i, 'size' => 11 + $i % 3]);
            $calculator->calculateCellWidth('text', $font, $base, false);
        }
        $memory = memory_get_usage();
        for ($i = 6000; $i < 12000; ++$i) {
            $font = $calculator->resolveCellFont($base, ['name' => 'Unknown ' . $i, 'size' => 11 + $i % 3]);
            $calculator->calculateCellWidth('text', $font, $base, false);
        }
        $this->assertLessThan(1048576, memory_get_usage() - $memory);
    }

    public function testWorkbookInheritanceContextsAreIndependent(): void
    {
        $first = Excel::create();
        $second = Excel::create();
        $base = ['name' => 'Calibri', 'size' => 11];
        $this->assertSame($base, $first->getAutoWidthFont()->resolveCellFont($base, []));
        $this->assertSame(['name' => 'Arial', 'size' => 11], $second->getAutoWidthFont()->resolveCellFont($base, ['name' => 'Arial']));
        $this->assertSame($base, $first->getAutoWidthFont()->resolveCellFont($base, []));
    }

    public function testCallbacksCanChangeTheBaseFontAndWidthFactor(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setColWidthAuto('A')->setColDataStyle('A', ['font-name' => 'Calibri', 'font-size' => 12]);
        $expected = null;
        $calls = 0;
        $excel->setAutoWidthFormatter(function () use ($excel, &$expected, &$calls) {
            ++$calls;
            $excel->setDefaultFont(['font-name' => 'Arial', 'font-size' => 16]);
            Font::setFontWidthFactor(1.2);
            $expected = (new Font())->calculateCellWidth('Text from callback', ['name' => 'Calibri', 'size' => 12],
                $excel->styleManager->defaultFont['val'], false);
            return 'Text from callback';
        });
        $sheet->writeRow(['original text']);
        $path = tempnam(sys_get_temp_dir(), 'few-font-');
        try {
            $excel->save($path);
            $zip = new ZipArchive();
            $zip->open($path);
            $source = $zip->getFromName('xl/worksheets/sheet1.xml');
            $zip->close();
            $xml = simplexml_load_string($source);
            $this->assertSame(1, $calls);
            $this->assertEqualsWithDelta($expected, (float)(string)$xml->cols->col['width'], 1 / 256);
        }
        finally {
            unlink($path);
        }
    }
}
