<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use PHPUnit\Framework\TestCase;

final class BufferedRowsTest extends TestCase
{
    public function futureCellModes(): array
    {
        return array_map(static function ($mode) {
            return [$mode];
        }, [
            'value', 'date', 'callable', 'styled-null', 'cell-style',
            'range-style', 'range-fill', 'border', 'image',
        ]);
    }

    /** @dataProvider futureCellModes */
    public function testFutureCellsSurvivePartialFlushes(string $mode): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        switch ($mode) {
            case 'value':
                $sheet->setValue('C9', 99);
                break;
            case 'date':
                $sheet->setValue('C9', new DateTimeImmutable('2026-10-09'));
                break;
            case 'callable':
                $sheet->setValue('C9', static function () { return 99; });
                break;
            case 'styled-null':
                $sheet->setValue('C9', null, ['fill-color' => '#123456']);
                break;
            case 'cell-style':
                $sheet->setCellStyle('C9', ['fill-color' => '#123456']);
                break;
            case 'range-style':
                $sheet->setCellStyle('B7:C9', ['fill-color' => '#123456']);
                break;
            case 'range-fill':
                $sheet->withRange('B7:C9')->applyBgColor('#123456');
                break;
            case 'border':
                $sheet->withRange('C9')->applyBorderBottom('dotted');
                break;
            case 'image':
                $sheet->addImage('C9', __DIR__ . '/../demo/logo/excel-logo.png');
                break;
        }
        // Insert a lower row after the highest row: insertion order is not row order.
        $sheet->setValue('B5', 'middle');
        $sheet->setTopLeftCell('A1')->writeRow(['first']);
        $sheet->writeRow(['second']);
        $this->assertSame(1, $sheet->rowCountWritten);
        $sheet->setTopLeftCell('A7')->writeRow(['seventh']);
        $this->assertSame(6, $sheet->rowCountWritten);

        $this->withSavedSheet($excel, function ($xml, $zip) use ($mode) {
            $this->assertSame('first', (string)$xml->xpath('//s:c[@r="A1"]/s:is/s:t')[0]);
            $this->assertSame('second', (string)$xml->xpath('//s:c[@r="A2"]/s:is/s:t')[0]);
            $this->assertSame('middle', (string)$xml->xpath('//s:c[@r="B5"]/s:is/s:t')[0]);
            $this->assertSame('seventh', (string)$xml->xpath('//s:c[@r="A7"]/s:is/s:t')[0]);
            $cells = $xml->xpath('//s:c[@r="C9"]');
            if ($mode === 'image') {
                $this->assertNotFalse($zip->locateName('xl/drawings/drawing1.xml'));
                $drawing = simplexml_load_string($zip->getFromName('xl/drawings/drawing1.xml'));
                $drawing->registerXPathNamespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
                $this->assertSame('8', (string)$drawing->xpath('//xdr:from/xdr:row')[0]);
                return;
            }
            $this->assertCount(1, $cells);
            if ($mode === 'value' || $mode === 'callable') {
                $this->assertSame('99', (string)$cells[0]->v);
            } elseif ($mode === 'date') {
                $this->assertNotEmpty((string)$cells[0]->v);
            } else {
                $this->assertGreaterThan(0, (int)$cells[0]['s']);
            }
            $this->assertSame('9', (string)$xml->sheetData->row[count($xml->sheetData->row) - 1]['r']);
        });
        $this->assertSame(9, $sheet->rowCountWritten);
    }

    public function testEmptyBufferCanBeRefilledAfterAreaFlush(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->makeArea('A1:C3')->setValue('C3', 'area');
        $sheet->writeAreas();
        $this->assertSame(3, $sheet->rowCountWritten);
        $sheet->withRange('B8')->applyBorderBottom('dotted');
        $sheet->setTopLeftCell('A4')->writeRow(['after area']);
        $sheet->setTopLeftCell('A9')->writeRow(['after border']);
        $this->assertSame(8, $sheet->rowCountWritten);
        $sheet->setTopLeftCell('A10')->writeRow(['refilled']);

        $this->withSavedSheet($excel, function ($xml) {
            foreach (['C3' => 'area', 'A4' => 'after area', 'A9' => 'after border', 'A10' => 'refilled'] as $address => $value) {
                $this->assertSame($value, (string)$xml->xpath('//s:c[@r="' . $address . '"]/s:is/s:t')[0]);
            }
            $this->assertCount(1, $xml->xpath('//s:c[@r="B8"]'));
            $this->assertSame('10', (string)$xml->sheetData->row[count($xml->sheetData->row) - 1]['r']);
        });
    }

    public function testAreaFlushIncludesStyleOnlyFutureRows(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setValue('A1', 'first');
        $sheet->setValue('C9', null, ['fill-color' => '#123456']);
        $sheet->makeArea('A1:C9');
        $sheet->writeAreas();
        $this->assertSame(9, $sheet->rowCountWritten);

        $this->withSavedSheet($excel, function ($xml) {
            $cells = $xml->xpath('//s:c[@r="C9"]');
            $this->assertCount(1, $cells);
            $this->assertGreaterThan(0, (int)$cells[0]['s']);
        });
    }

    public function testFutureRowAttributesAreStillSaved(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setRowHeight(12, 30);
        $sheet->withRange('B9')->applyBgColor('#123456');
        $sheet->writeRow(['first'])->writeRow(['second']);

        $this->withSavedSheet($excel, function ($xml) {
            $rows = $xml->xpath('//s:row[@r="12"]');
            $this->assertCount(1, $rows);
            $this->assertSame('30', (string)$rows[0]['ht']);
            $this->assertCount(1, $xml->xpath('//s:c[@r="B9"]'));
        });
    }

    public function testValueRemainsBufferedWhenStyleNormalizationFails(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        try {
            $sheet->setValue('C9', 99, ['font-style' => new stdClass()]);
            $this->fail('Invalid font style must throw an error');
        } catch (\Error $error) {
            $this->assertStringContainsString('could not be converted to string', $error->getMessage());
        }
        $sheet->setTopLeftCell('A1')->writeRow(['first']);
        $this->withSavedSheet($excel, function ($xml) {
            $this->assertSame('99', (string)$xml->xpath('//s:c[@r="C9"]/s:v')[0]);
            $this->assertSame('first', (string)$xml->xpath('//s:c[@r="A1"]/s:is/s:t')[0]);
        });
    }

    private function withSavedSheet(Excel $excel, callable $check): void
    {
        $file = tempnam(sys_get_temp_dir(), 'buffered-rows-');
        $zip = new ZipArchive();
        try {
            $excel->save($file);
            $this->assertTrue($zip->open($file) === true);
            $xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
            $xml->registerXPathNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
            $check($xml, $zip);
        } finally {
            $zip->close();
            unlink($file);
        }
    }
}
