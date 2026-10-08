<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\Style;
use PHPUnit\Framework\TestCase;

final class BorderMethodsTest extends TestCase
{
    public function borderCases(): array
    {
        $cases = [];
        foreach (['B2' => [2, 2], 'B2:D2' => [2, 4], 'B2:B4' => [4, 2], 'B2:D4' => [4, 4]] as $range => $limits) {
            foreach ([false, true] as $area) {
                foreach ([false, true] as $aliases) {
                    foreach (['#f00', 'none'] as $color) {
                        $name = $range . '-' . ($area ? 'area' : 'sheet') . '-' . ($aliases ? 'aliases' : 'primary') . '-' . $color;
                        $cases[$name] = [$range, $limits, $area, $aliases, $color];
                    }
                }
            }
        }
        return $cases;
    }

    /** @dataProvider borderCases */
    public function testRangeBorderMethods(string $range, array $limits, bool $area, bool $aliases, string $color): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->writeTo('A1', 1);
        $target = $area ? $sheet->makeArea($range) : $sheet;
        $target->withRange($range);
        $outer = $aliases ? 'applyOuterBorder' : 'applyBorderOuter';
        $inner = $aliases ? 'applyInnerBorder' : 'applyBorderInner';
        $this->assertSame($target, $target->$outer(Style::BORDER_DOUBLE, $color));
        $this->assertSame($target, $target->$inner(Style::BORDER_DOTTED, $color));
        for ($row = 1; $row <= 4; $row++) {
            $sheet->writeRow([$row * 10 + 1, $row * 10 + 2, $row * 10 + 3, $row * 10 + 4]);
        }
        $file = tempnam(sys_get_temp_dir(), 'border-methods-');
        $zip = new ZipArchive();
        $opened = false;
        try {
            $excel->save($file);
            $opened = $zip->open($file) === true;
            $this->assertTrue($opened);
            $styles = simplexml_load_string($zip->getFromName('xl/styles.xml'));
            $xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
            $xml->registerXPathNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
            for ($row = 2; $row <= $limits[0]; $row++) {
                for ($col = 2; $col <= $limits[1]; $col++) {
                    $address = chr(64 + $col) . $row;
                    $cells = $xml->xpath('//s:c[@r="' . $address . '"]');
                    $this->assertCount(1, $cells, $address);
                    if (!$area) {
                        $this->assertSame((string)(($row - 1) * 10 + $col), (string)$cells[0]->v, $address . ':value');
                    }
                    $xf = $styles->cellXfs->xf[(int)$cells[0]['s']];
                    $border = $styles->borders->border[(int)$xf['borderId']];
                    $expected = [
                        'left' => $col === 2 ? 'double' : '',
                        'top' => $row === 2 ? 'double' : '',
                        'right' => $col === $limits[1] ? 'double' : 'dotted',
                        'bottom' => $row === $limits[0] ? 'double' : 'dotted',
                    ];
                    foreach ($expected as $side => $style) {
                        $this->assertSame($style, (string)$border->$side['style'], $address . ':' . $side);
                        if ($style !== '') {
                            $attribute = $color === 'none' ? 'auto' : 'rgb';
                            $value = $color === 'none' ? '1' : 'FFFF0000';
                            $this->assertSame($value, (string)$border->$side->color[$attribute], $address . ':' . $side);
                        }
                    }
                }
            }
            // Border application must not affect cells outside the selection.
            $cells = $xml->xpath('//s:c[@r="A1"]');
            $this->assertCount(1, $cells);
            $this->assertSame(0, (int)$styles->cellXfs->xf[(int)$cells[0]['s']]['borderId']);
        }
        finally {
            if ($opened) {
                $zip->close();
            }
            unlink($file);
        }
    }
}
