<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\DisplayFormatter;
use avadim\FastExcelWriter\Writer\Writer;
use avadim\FastExcelReader\Excel as ExcelReader;
use PHPUnit\Framework\TestCase;

class CountingDateFormatter extends DisplayFormatter
{
    public int $calls = 0;

    protected function date(float $serial, array $tokens, string $locale): ?string
    {
        ++$this->calls;
        return parent::date($serial, $tokens, $locale);
    }
}

class UncachedDateFormatter extends DisplayFormatter
{
    protected function formatDate(float $serial, array $tokens, string $locale, string $format): ?string
    {
        return $this->date($serial, $tokens, $locale);
    }
}

final class AutoWidthDateFormattingTest extends TestCase
{
    public function testIdenticalDaysOnlyRenderOnceAndOtherTypesDoNotReuseThem(): void
    {
        $formatter = new CountingDateFormatter();
        $mask = 'dddd, d mmmm yyyy';
        for ($i = 0; $i < 24; ++$i) {
            $this->assertSame('Monday, 1 January 2024', $formatter->format(45292, $mask, true, 'en'));
        }
        $this->assertSame(1, $formatter->calls);
        $this->assertSame('45292', $formatter->format(45292, $mask, false, 'en'));
        $this->assertSame('TRUE', $formatter->format(true, $mask, true, 'en'));
        $this->assertSame('', $formatter->format(null, $mask, true, 'en'));
        $this->assertSame('Tuesday, 2 January 2024', $formatter->format(45293, $mask, true, 'en'));
        $this->assertSame(2, $formatter->calls);
    }

    public function testMasksLocalesFractionsAndConditionsMatchTheUncachedFormatter(): void
    {
        $cached = new DisplayFormatter();
        $reference = new UncachedDateFormatter();
        $values = [0, 1, 59, 60, 60.5, 61, 61.25, 61.75, 61.9999999999, 45292, 45292.5,
            2958465, -1, 2958466, '00060', '-01', true, false, null, INF, NAN];
        $formats = ['dddd, d mmmm yyyy', 'ddd mmm yy', 'mmmmm', 'yyyy-mm-dd hh:mm:ss', 'h:mm AM/PM',
            '[h]:mm:ss', '[$-419]dddd, d mmmm yyyy', '[>=61]dddd;"early "mmmm', '0;0;0;"text "@'];
        for ($pass = 0; $pass < 2; ++$pass) {
            foreach ($values as $value) {
                foreach (['en', 'ru', 'de', 'fr'] as $locale) {
                    foreach ($formats as $format) {
                        foreach ([true, false] as $numeric) {
                            $this->assertSame($reference->format($value, $format, $numeric, $locale),
                                $cached->format($value, $format, $numeric, $locale));
                        }
                    }
                }
            }
        }
        $this->assertSame('-01', $cached->format('-01', 'dddd', true));
        $this->assertSame('-1', $cached->format(-1, 'dddd', true));
    }

    public function testEffectiveLocaleAndMaskChangesInvalidateTheLastResult(): void
    {
        $formatter = new CountingDateFormatter();
        $mask = '[$-419]dddd, d mmmm yyyy';
        $this->assertSame('понедельник, 1 январь 2024', $formatter->format(45292, $mask, true, 'en'));
        $this->assertSame('понедельник, 1 январь 2024', $formatter->format(45292, $mask, true, 'de'));
        $this->assertSame(1, $formatter->calls);
        $this->assertSame('Monday', $formatter->format(45292, 'dddd', true, 'en'));
        $this->assertSame('понедельник', $formatter->format(45292, 'dddd', true, 'ru'));
        $this->assertSame(3, $formatter->calls);
    }

    public function testOversizedMasksLocalesAndOutputAreNotRetained(): void
    {
        $formatter = new CountingDateFormatter();
        foreach ([['yyyy', str_repeat('e', 257), 45292], [str_repeat('y', 4097), 'en', 45292],
            [str_repeat('d ', 1500), 'en', 45322]] as [$mask, $locale, $serial]) {
            $before = $formatter->calls;
            $formatter->format($serial, $mask, true, $locale);
            $formatter->format($serial, $mask, true, $locale);
            $this->assertSame($before + 2, $formatter->calls);
        }
    }

    public function testRepeatedDateWidthsStillCallTheCallbackForEveryCell(): void
    {
        $excel = Excel::create([], ['locale' => 'en']);
        $sheet = $excel->sheet()->setColWidthAuto('A')->setColDataStyle('A', ['format' => 'dddd, d mmmm yyyy']);
        $expectedFormat = $excel->styleManager->defineFormatType('dddd, d mmmm yyyy')['number_format'];
        $seen = [];
        $excel->setAutoWidthFormatter(function ($value, $format, $type, $locale) use (&$seen) {
            $seen[] = [$value, $format, $type, $locale];
            return null;
        });
        $source = 1704067200;
        for ($i = 0; $i < 24; ++$i) {
            $sheet->writeRow([$source + $i * 3600]);
        }
        $path = tempnam(sys_get_temp_dir(), 'few-date-format-');
        try {
            $excel->save($path);
            $this->assertTrue(ExcelReader::validate($path, $errors));
            $this->assertCount(24, $seen);
            foreach ($seen as $index => $arguments) {
                $this->assertSame([(int)Writer::convertDateTime($source + $index * 3600), $expectedFormat, 'n_date', 'en'], $arguments);
            }
        }
        finally {
            unlink($path);
        }
    }

    public function testUniqueDaysAndMasksDoNotKeepGrowingMemory(): void
    {
        $formatter = new DisplayFormatter();
        for ($i = 0; $i < 1000; ++$i) {
            $formatter->format(45292 + $i, 'dddd, d mmmm yyyy "' . $i . '"', true, 'en');
        }
        $memory = memory_get_usage();
        for ($i = 1000; $i < 5000; ++$i) {
            $formatter->format(45292 + $i, 'dddd, d mmmm yyyy "' . $i . '"', true, 'en');
        }
        $this->assertLessThan(1048576, memory_get_usage() - $memory);
    }
}
