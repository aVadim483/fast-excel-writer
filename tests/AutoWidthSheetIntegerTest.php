<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\Font;
use PHPUnit\Framework\TestCase;

final class AutoWidthSheetIntegerTest extends TestCase
{
    protected function tearDown(): void { Font::setFontWidthFactor(1.0); }

    protected function measure(Excel $excel, $value, string $format = 'GENERAL', array $style = []): ?float
    {
        return Closure::bind(function ($value, $format, $style) {
            $this->_columnWidth(0, $value, $format, $style, false);
            return $this->colAttributes[0]['width'] ?? null;
        }, $excel->sheet(), get_class($excel->sheet()))($value, $format, $style);
    }

    public function testNativeIntegersPreserveFontsFormatsFilterAndScaling(): void
    {
        foreach ([0.0, 1.0, 1.5] as $factor) {
            Font::setFontWidthFactor($factor);
            foreach ([false, true] as $filter) {
                foreach (['GENERAL', '@', '0'] as $format) {
                    foreach ([0, 12, -12, 999999999999999] as $value) {
                        $excel = Excel::create();
                        $excel->setDefaultFont(['font-name' => 'Tahoma', 'font-size' => 14]);
                        if ($filter) { $excel->sheet()->setAutoFilter('A1:A10'); }
                        $base = $excel->styleManager->defaultFont['val'];
                        $override = ['name' => 'Arial', 'size' => 16, 'style-bold' => true];
                        $font = array_replace($base, $override);
                        $expected = (new Font())->calculateCellWidth((string)$value, $font, $base, false);
                        $this->assertSame(min(255.0, $expected + ($filter ? 1 : 0)),
                            $this->measure($excel, $value, $format, ['font' => ['val' => $override]]));
                    }
                }
            }
        }
    }

    public function testColumnWidthsOnlyGrowAndRespectMinimumAndMaximum(): void
    {
        $excel = Excel::create();
        $excel->sheet()->setColMinWidth('A', 10);
        $this->assertSame(10.0, $this->measure($excel, 12));
        $this->assertSame(10.0, $this->measure($excel, 0));
        $base = $excel->styleManager->defaultFont['val'];
        $expected = (new Font())->calculateCellWidth((string)PHP_INT_MAX, $base, $base, false);
        $this->assertSame($expected, $this->measure($excel, PHP_INT_MAX));
        $this->assertSame($expected, $this->measure($excel, 1));
        Font::setFontWidthFactor(1000.0);
        $excel->sheet()->setAutoFilter('A1:A10');
        $this->assertSame(255.0, $this->measure($excel, PHP_INT_MAX));
        $this->assertSame(255.0, $this->measure($excel, 0));
    }

    public function testFormattedAndWrappedValuesKeepTheirExistingPaths(): void
    {
        foreach ([[12, '000000', '000012'], [12, '0%', '1200%'], [PHP_INT_MAX, '0', null],
            ['000012', '@', '000012'], [['=1+1', 2], '0.00', '2.00'], [['value' => 12], 'GENERAL', '12']] as [$value, $format, $text]) {
            $excel = Excel::create();
            if ($text === null) { $text = $excel->formatAutoWidthValue($value, $format, 'n_numeric', true); }
            $base = $excel->styleManager->defaultFont['val'];
            $expected = (new Font())->calculateCellWidth($text, $base, $base, false);
            $this->assertSame($expected, $this->measure($excel, $value, $format));
        }
    }

    public function testCallbacksStillRunAndCanChangeTheColumnWidth(): void
    {
        $excel = Excel::create(); $seen = [];
        $excel->setAutoWidthFormatter(function ($value) use ($excel, &$seen) {
            $seen[] = $value;
            $excel->sheet()->setColWidth('A', 100);
            return 'x';
        });
        $this->assertSame(100.0, $this->measure($excel, 12));
        $this->assertSame(100.0, $this->measure($excel, 0));
        $excel->sheet()->setColWidth('A', 255);
        $this->measure($excel, 34);
        $this->assertSame([12, 0, 34], $seen);
    }
}
