<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Style\DisplayFormatter;
use PHPUnit\Framework\TestCase;

class NativeDatePlanFormatter extends DisplayFormatter
{
    public int $builds = 0;
    public function nativeMask(string $format): ?string
    {
        return $this->formatPlan($format)['sections'][0]['date_plan']['native_mask'];
    }
    protected function buildCalendarParts(int $day): array
    {
        ++$this->builds;
        return parent::buildCalendarParts($day);
    }
}

class OperationDatePlanFormatter extends DisplayFormatter
{
    protected function prepareDate(array $tokens): array
    {
        $plan = parent::prepareDate($tokens);
        $plan['native_mask'] = null;
        return $plan;
    }
}

final class AutoWidthNativeDateTest extends TestCase
{
    public function testCompiledCalendarMasksMatchTheOperationRenderer(): void
    {
        $native = new DisplayFormatter();
        $reference = new OperationDatePlanFormatter();
        $formats = ['dddd, d mmmm yyyy', 'ddd, dd mmm yy', 'yyyy-mm-dd hh:mm:ss',
            'yyyy-mm-dd h:mm AM/PM', 'yyyy-mm-dd h:mm "AM/PM"', 'yyy mmmmmm ddddd',
            'mm hh mm ss yyyy', 'mmmm "x" mm:ss', 'yyyy m:s', 'yyyy hhh:mm:ss', 'mmmmm yyyy',
            'yyyy h:mm A/P', 'yyyy [h]:mm:ss', '"Ymd H:i:s \\ русский 🙂" dddd yyyy', "yyyy\nmmmm dd",
            'yyyy "' . "\0" . '" mmmm', '[$-419]dddd, d mmmm yyyy', '[$-409]dddd, d mmmm yyyy',
            '[>=61]dddd, d mmmm yyyy;"early "mmmm', 'yyyy;"negative "dddd;"zero "mmmm;@'];
        $values = [0, 0.5, 1, 59, 59.9999999999, 60, 60.5, 60.9999999999, 61, 61.75, 45292,
            45292 + 47107 / 86400, 45292.9999999999, 2958465, 2958465.99999999, -1, 2958466, '00060', INF, NAN];
        $timezone = date_default_timezone_get();
        try {
            date_default_timezone_set('Pacific/Honolulu');
            foreach (['en', 'ru', 'en-US', 'de'] as $locale) {
                foreach ($formats as $format) {
                    foreach ($values as $value) {
                        $this->assertSame($reference->format($value, $format, true, $locale),
                            $native->format($value, $format, true, $locale), $format . ' / ' . $locale);
                    }
                }
            }
        }
        finally { date_default_timezone_set($timezone); }
    }

    public function testNativeFormattingKeepsExcelSyntheticDaysAndEffectiveLocale(): void
    {
        $formatter = new NativeDatePlanFormatter();
        $mask = 'dddd, d mmmm yyyy';
        $this->assertSame('Monday, 1 January 2024', $formatter->format(45292, $mask, true));
        $this->assertSame('Tuesday, 2 January 2024', $formatter->format(45293, $mask, true));
        if (PHP_INT_SIZE >= 8) { $this->assertSame(0, $formatter->builds); }
        $before = $formatter->builds;
        $this->assertSame('Sunday, 0 January 1900', $formatter->format(0, $mask, true));
        $this->assertSame('Wednesday, 29 February 1900', $formatter->format(60, $mask, true));
        $this->assertSame($before + 2, $formatter->builds);
        $this->assertSame('понедельник, 1 январь 2024', $formatter->format(45292, '[$-419]' . $mask, true));
        $this->assertSame('Monday, 1 January 2024', $formatter->format(45292, '[$-409]' . $mask, true, 'ru'));
        $this->assertSame('2024-01-02 00:00:00', $formatter->format(45292.9999999999, 'yyyy-mm-dd hh:mm:ss', true));
        $this->assertSame('10000-01-01', $formatter->format(2958465.99999999, 'yyyy-mm-dd', true));
    }

    public function testUnsupportedAndTimeOnlyMasksKeepTheOperationRenderer(): void
    {
        $formatter = new NativeDatePlanFormatter();
        $this->assertNotNull($formatter->nativeMask('dddd, d mmmm yyyy'));
        $this->assertSame($formatter->nativeMask('dddd, d mmmm yyyy'), $formatter->nativeMask('""dddd, d mmmm yyyy'));
        $this->assertSame($formatter->nativeMask('dddd, d mmmm yyyy'), $formatter->nativeMask('[$-409]dddd, d mmmm yyyy'));
        foreach (['mmmmm', 'mmmmmm', 'yyyy hhh:mm:ss', 'yyyy m:s', 'yyyy h:mm A/P',
            'yyyy [h]:mm:ss', 'hh:mm:ss', 'yyyy "' . "\0" . '" mmmm'] as $format) {
            $this->assertNull($formatter->nativeMask($format), $format);
        }
        $this->assertSame('Monday, 1 January 2024', $formatter->format(45292, 'dddd, d mmmm yyyy', true));
    }
}
