<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Style\Font;
use PHPUnit\Framework\TestCase;

class CountingIntegerWidthFont extends Font
{
    public int $conversions = 0;

    protected function excelWidth(float $pixels, array $baseFont): float
    {
        ++$this->conversions;
        return parent::excelWidth($pixels, $baseFont);
    }

    public function cachedWidths(): int { return count($this->simpleWidths) + count($this->previousSimpleWidths); }
}

final class AutoWidthIntegerCacheTest extends TestCase
{
    protected function tearDown(): void { Font::setFontWidthFactor(1.0); }

    public function testRepeatedLengthsReuseReadyWidthsAndInvalidateTheScale(): void
    {
        $font = ['name' => 'Calibri', 'size' => 11];
        $calculator = new CountingIntegerWidthFont();
        foreach ([12, 34, 99, 12, -12, -34, -99] as $value) {
            $this->assertSame((new Font())->calculateCellWidth((string)$value, $font, [], false),
                $calculator->calculateIntegerWidth($value, $font, []));
        }
        $this->assertSame(2, $calculator->conversions);
        foreach ([0.0, 0.8, 1.0, 1000.0, 1.0] as $factor) {
            Font::setFontWidthFactor($factor);
            $before = $calculator->conversions;
            foreach ([12, 34, 99] as $value) {
                $this->assertSame((new Font())->calculateCellWidth((string)$value, $font, [], false),
                    $calculator->calculateIntegerWidth($value, $font, []));
            }
            $this->assertSame($before + 1, $calculator->conversions);
        }
    }

    public function testFontAndBaseFontChangesMatchFreshCalculators(): void
    {
        $calculator = new Font();
        foreach (['Calibri', 'Arial', 'Tahoma', 'Times New Roman', 'Unknown font'] as $name) {
            foreach ([11, 16] as $size) {
                foreach ([0, 1, 2, 3] as $face) {
                    $font = ['name' => $name, 'size' => $size, 'style-bold' => (bool)($face & 1), 'style-italic' => (bool)($face & 2)];
                    foreach ([[], ['name' => 'Tahoma', 'size' => 14]] as $base) {
                        foreach ([0, 1, 9, 10, 99, -1, -9, -10, -99, PHP_INT_MIN, PHP_INT_MAX] as $value) {
                            $this->assertSame((new Font())->calculateCellWidth((string)$value, $font, $base, false),
                                $calculator->calculateIntegerWidth($value, $font, $base));
                        }
                    }
                }
            }
        }
    }

    public function testAlternatingFontsReuseWidthsAndBothInvalidateWhenTheScaleChanges(): void
    {
        $calculator = new CountingIntegerWidthFont();
        $fonts = [['name' => 'Calibri', 'size' => 11], ['name' => 'Arial', 'size' => 16]];
        foreach ([[1.0, []], [1.5, []], [1.5, ['name' => 'Tahoma', 'size' => 14]]] as [$factor, $base]) {
            Font::setFontWidthFactor($factor);
            $before = $calculator->conversions;
            for ($i = 0; $i < 20; ++$i) {
                $font = $fonts[$i % 2];
                $this->assertSame((new Font())->calculateCellWidth('12', $font, $base, false),
                    $calculator->calculateIntegerWidth(12, $font, $base));
            }
            $this->assertSame($before + 2, $calculator->conversions);
        }
    }

    public function testReferencedFontPropertiesAreSnapshotted(): void
    {
        $calculator = new Font();
        $name = 'Calibri'; $size = 11; $bold = false; $baseSize = 11;
        $font = ['name' => &$name, 'size' => &$size, 'style-bold' => &$bold];
        $base = ['name' => 'Arial', 'size' => &$baseSize];
        foreach ([['Calibri', 11, false, 11], ['Arial', 16, true, 11], ['Arial', 16, true, 22],
            ['Tahoma', 12, false, 14], ['Calibri', 11, false, 11]] as $properties) {
            [$name, $size, $bold, $baseSize] = $properties;
            foreach ([12, 34, -12, 0, PHP_INT_MIN] as $value) {
                $expected = (new Font())->calculateCellWidth((string)$value,
                    ['name' => $name, 'size' => $size, 'style-bold' => $bold], ['name' => 'Arial', 'size' => $baseSize], false);
                $this->assertSame($expected, $calculator->calculateIntegerWidth($value, $font, $base));
            }
        }
    }

    public function testTheCacheIsBoundedByLengthsAndBelongsToEachCalculator(): void
    {
        $first = new CountingIntegerWidthFont(); $second = new CountingIntegerWidthFont();
        for ($value = -10000; $value <= 10000; ++$value) { $first->calculateIntegerWidth($value, [], []); }
        $first->calculateIntegerWidth(PHP_INT_MIN, [], []);
        $first->calculateIntegerWidth(PHP_INT_MAX, [], []);
        $this->assertLessThanOrEqual(2 * strlen((string)PHP_INT_MIN), $first->cachedWidths());
        $this->assertSame(0, $second->cachedWidths());
        $second->calculateIntegerWidth(12, [], []);
        $this->assertSame(1, $second->cachedWidths());
    }
}
