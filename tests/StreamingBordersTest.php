<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\Style;
use PHPUnit\Framework\TestCase;

final class StreamingBordersTest extends TestCase
{
    public function writingModes(): array
    {
        return [['before'], ['preformat'], ['position'], ['after'], ['flat'], ['nested']];
    }

    /** @dataProvider writingModes */
    public function testValuesAndBordersSurviveStreaming(string $mode): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->writeHeader(['Name', 'Amount']);
        $data = [['James', 220], ['Mike', 153.5], ['John', 34.12]];
        $border = ['border-bottom-style' => 'dotted', 'border-bottom-color' => '#123456'];
        if ($mode === 'preformat') {
            $sheet->withRange('A2:B4')->applyBorderBottom('dotted', '#123456');
            // Bottom borders are applied to each row, not just the range outline.
            $sheet->withRange('A2:B2')->applyBorderBottom('dotted', '#123456');
            $sheet->withRange('A3:B3')->applyBorderBottom('dotted', '#123456');
            $sheet->setTopLeftCell('A2');
        }
        foreach ($data as $index => $values) {
            $row = $index + 2;
            if ($mode === 'before' || $mode === 'position') {
                $sheet->withRange('A' . $row . ':B' . $row)->applyBorderBottom('dotted', '#123456');
            }
            if ($mode === 'position') {
                $sheet->setTopLeftCell('A' . $row);
            }
            $sheet->writeRow($values, $mode === 'flat' ? $border : ($mode === 'nested' ? ['border' => $border] : null));
            if ($mode === 'after') {
                $sheet->applyBorderBottom('dotted', '#123456');
            }
            $this->assertSame($row - 1, $sheet->rowCountWritten, 'Only preceding rows should be flushed');
        }
        $file = tempnam(sys_get_temp_dir(), 'streaming-borders-');
        $zip = new ZipArchive();
        try {
            $excel->save($file);
            $this->assertTrue($zip->open($file) === true);
            $xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
            $styles = simplexml_load_string($zip->getFromName('xl/styles.xml'));
            $this->assertCount(4, $xml->sheetData->row);
            foreach ($data as $index => $values) {
                $row = $xml->sheetData->row[$index + 1];
                $this->assertSame((string)($index + 2), (string)$row['r']);
                $this->assertSame($values[0], (string)$row->c[0]->is->t);
                $this->assertSame((string)$values[1], (string)$row->c[1]->v);
                foreach ($row->c as $cell) {
                    $borderId = (int)$styles->cellXfs->xf[(int)$cell['s']]['borderId'];
                    $bottom = $styles->borders->border[$borderId]->bottom;
                    $this->assertSame('dotted', (string)$bottom['style']);
                    $this->assertSame('FF123456', (string)$bottom->color['rgb']);
                }
            }
        } finally {
            $zip->close();
            unlink($file);
        }
    }

    public function testExplicitPositionKeepsOnlyTheCurrentRowBuffered(): void
    {
        $sheet = Excel::create()->sheet();
        $property = new ReflectionProperty($sheet, 'cells');
        $property->setAccessible(true);
        for ($row = 1; $row <= 100; $row++) {
            $sheet->setTopLeftCell('C' . $row)->writeRow([$row]);
            $this->assertSame($row - 1, $sheet->rowCountWritten);
            $buffer = $property->getValue($sheet);
            $this->assertSame([$row - 1], array_keys($buffer['values']));
        }
        $this->expectException(\avadim\FastExcelWriter\Exceptions\Exception::class);
        $sheet->withRange('C1')->applyBorderBottom('thin');
    }

    public function testSaveIncludesFutureStyledRows(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->writeRow(['first']);
        $sheet->withRange('A5')->applyBorderBottom('dotted');
        $sheet->writeRow(['second']);
        $this->assertSame(1, $sheet->rowCountWritten);
        $file = tempnam(sys_get_temp_dir(), 'future-border-');
        $zip = new ZipArchive();
        try {
            $excel->save($file);
            $this->assertTrue($zip->open($file) === true);
            $xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
            $xml->registerXPathNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
            $this->assertCount(1, $xml->xpath('//s:c[@r="A5"]'));
            $this->assertSame('second', (string)$xml->sheetData->row[1]->c->is->t);
        } finally {
            $zip->close();
            unlink($file);
        }
    }
}
