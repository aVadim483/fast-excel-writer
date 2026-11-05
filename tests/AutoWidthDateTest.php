<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\RichText\RichText;
use avadim\FastExcelWriter\Writer\FileWriter;
use avadim\FastExcelWriter\Writer\PreparedDateTime;
use avadim\FastExcelWriter\Writer\Writer;
use avadim\FastExcelReader\Excel as ExcelReader;
use PHPUnit\Framework\TestCase;

class AutoWidthDateWriter extends Writer
{
    public array $preparedCells = [];
    public array $seenValues = [];
    public ?string $replacement = null;
    public bool $fail = false;

    public function _writePreparedDateCell(FileWriter $file, int $rowNumber, int $colNumber, $value,
        $numFormatType, ?int $cellStyleIdx, PreparedDateTime $date)
    {
        $this->preparedCells[] = [$rowNumber, $colNumber, $date->sourceValue, $date->serial];
        parent::_writePreparedDateCell($file, $rowNumber, $colNumber, $value, $numFormatType, $cellStyleIdx, $date);
    }

    public function _writeCell(FileWriter $file, int $rowNumber, int $colNumber, $value, $numFormatType, ?int $cellStyleIdx = 0)
    {
        $this->seenValues[] = $value;
        if ($this->fail) {
            throw new RuntimeException('Cell write failed');
        }
        if ($this->replacement !== null) {
            $value = $this->replacement;
        }
        parent::_writeCell($file, $rowNumber, $colNumber, $value, $numFormatType, $cellStyleIdx);
    }

    public function hasPreparedDate(): bool
    {
        return $this->preparedDateTime !== null;
    }
}

final class AutoWidthDateTest extends TestCase
{
    protected array $files = [];

    protected function tearDown(): void
    {
        foreach ($this->files as $file) {
            if (is_file($file)) {
                unlink($file);
            }
        }
    }

    protected function tempFile(): string
    {
        $file = tempnam(sys_get_temp_dir(), 'few-date-');
        $this->files[] = $file;
        return $file;
    }

    protected function saveCells(Excel $excel): array
    {
        $file = $this->tempFile();
        $excel->save($file);
        $this->assertTrue(ExcelReader::validate($file, $errors));
        $zip = new ZipArchive();
        $zip->open($file);
        $xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
        $zip->close();
        $cells = [];
        foreach ($xml->sheetData->row as $row) {
            foreach ($row->c as $cell) {
                $cells[(string)$cell['r']] = $cell;
            }
        }
        return $cells;
    }

    public function testDatesKeepTheSameXmlWithAndWithoutAutoWidth(): void
    {
        $excel = Excel::create([], ['writer_class' => AutoWidthDateWriter::class]);
        $sheet = $excel->sheet()->setColWidthAuto('A');
        $cases = [
            ['2024-01-01 12:34:56', '@datetime'],
            [1704067200, '@datetime'],
            [new DateTimeImmutable('2024-02-29T12:34:56+03:00'), '@datetime'],
            ['1900-02-29 12:00:00', '@datetime'],
            ['1900-02-29', '@date'],
            ['1898-05-05', '@date'],
            ['invalid date', '@datetime'],
            [0, '@datetime'],
            ['', '@date'],
            [null, '@date'],
            [false, '@date'],
            [true, '@date'],
            [new RichText('Rich <b>text</b>'), '@date'],
            [['=1+1', 2], '@date'],
        ];
        foreach ($cases as $index => [$value, $format]) {
            // Keep rows containing two null values in the streaming output.
            $sheet->writeRow([$value, $value, $index], ['format' => $format]);
        }
        $cells = $this->saveCells($excel);
        foreach ($cases as $index => $case) {
            $this->assertArrayHasKey('A' . ($index + 1), $cells);
            $this->assertArrayHasKey('B' . ($index + 1), $cells);
            $a = $cells['A' . ($index + 1)];
            $b = $cells['B' . ($index + 1)];
            $this->assertSame((string)$b['t'], (string)$a['t']);
            $this->assertSame((string)$b->v, (string)$a->v);
            $this->assertSame((string)$b->is->t, (string)$a->is->t);
            $this->assertSame((string)$b->f, (string)$a->f);
        }
        $writer = $excel->getWriter();
        $this->assertCount(8, $writer->preparedCells);
        $this->assertSame('2024-01-01 12:34:56', $writer->seenValues[0]);
        $this->assertSame(60.5, (float)(string)$cells['A4']->v);
        $this->assertSame(60.0, (float)(string)$cells['A5']->v);
        $this->assertSame('invalid date', (string)$cells['A7']->is->t);
        $this->assertSame('2', (string)$cells['A14']->v);
        $this->assertFalse($writer->hasPreparedDate());
    }

    public function testCallbackTimezoneChangesRetainTheirSerializationEffect(): void
    {
        $timezone = date_default_timezone_get();
        try {
            date_default_timezone_set('UTC');
            $source = 1704067200;
            $before = Writer::convertDateTime($source);
            $seen = [];
            $excel = Excel::create();
            $excel->setAutoWidthFormatter(function ($value, $format, $type) use (&$seen) {
                $seen[] = [$value, $type];
                date_default_timezone_set('Europe/Moscow');
                return null;
            });
            $excel->sheet()->setColWidthAuto('A')->writeRow([$source], ['format' => '@datetime']);
            $cells = $this->saveCells($excel);
            $this->assertSame([[$before, 'n_datetime']], $seen);
            $after = Writer::convertDateTime($source);
            $this->assertNotEquals($before, $after);
            $this->assertEqualsWithDelta($after, (float)(string)$cells['A1']->v, 1e-9);
        }
        finally {
            date_default_timezone_set($timezone);
        }
    }

    public function testCustomWriterCanChangeTheOriginalValue(): void
    {
        $excel = Excel::create([], ['writer_class' => AutoWidthDateWriter::class]);
        $writer = $excel->getWriter();
        $writer->replacement = '2025-05-06 12:00:00';
        $excel->sheet()->setColWidthAuto('A')->writeRow(['2024-01-01'], ['format' => '@datetime']);
        $cells = $this->saveCells($excel);
        $this->assertSame(['2024-01-01'], $writer->seenValues);
        $this->assertEqualsWithDelta(Writer::convertDateTime($writer->replacement), (float)(string)$cells['A1']->v, 1e-9);
        $this->assertFalse($writer->hasPreparedDate());
    }

    public function testPreparedDatesAreReusedAndReleasedAfterSuccessAndFailure(): void
    {
        $excel = Excel::create([], ['writer_class' => AutoWidthDateWriter::class]);
        $writer = $excel->getWriter();
        $path = $this->tempFile();
        $file = new FileWriter($path);
        $source = '2024-01-01 12:00:00';
        $date = new PreparedDateTime($source);
        $date->serial = 60.5;
        $writer->_writePreparedDateCell($file, 1, 1, $source, 'n_datetime', 0, $date);
        $this->assertFalse($writer->hasPreparedDate());
        $writer->_writeCell($file, 2, 1, $source, 'n_datetime');
        $writer->fail = true;
        try {
            $writer->_writePreparedDateCell($file, 3, 1, $source, 'n_datetime', 0, $date);
            $this->fail('The custom writer must throw');
        }
        catch (RuntimeException $error) {
            $this->assertSame('Cell write failed', $error->getMessage());
        }
        $this->assertFalse($writer->hasPreparedDate());
        $writer->fail = false;
        $writer->_writeCell($file, 4, 1, $source, 'n_datetime');
        $file->close();
        $xml = simplexml_load_string('<row>' . file_get_contents($path) . '</row>');
        $this->assertSame(60.5, (float)(string)$xml->c[0]->v);
        $expected = Writer::convertDateTime($source);
        $this->assertEqualsWithDelta($expected, (float)(string)$xml->c[1]->v, 1e-9);
        $this->assertEqualsWithDelta($expected, (float)(string)$xml->c[2]->v, 1e-9);
    }
}
