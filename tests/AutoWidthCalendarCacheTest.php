<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Style\DisplayFormatter;
use PHPUnit\Framework\TestCase;

class CountingCalendarFormatter extends DisplayFormatter
{
    public int $builds = 0;

    protected function buildCalendarParts(int $day): array
    {
        ++$this->builds;
        return parent::buildCalendarParts($day);
    }

    public function epoch(): ?DateTimeImmutable
    {
        return $this->dateEpoch;
    }
}

final class AutoWidthCalendarCacheTest extends TestCase
{
    public function testDifferentTimesAndLocalesReuseTheCalendarDay(): void
    {
        $formatter = new CountingCalendarFormatter();
        $mask = 'yyyy-mm-dd hh:mm:ss';
        for ($hour = 0; $hour < 24; ++$hour) {
            $this->assertSame('2024-01-01 ' . sprintf('%02d:00:00', $hour), $formatter->format(45292 + $hour / 24, $mask, true));
        }
        $this->assertSame('Monday, 1 January 2024', $formatter->format(45292.5, 'dddd, d mmmm yyyy', true, 'en'));
        $this->assertSame('понедельник, 1 январь 2024', $formatter->format(45292.5, 'dddd, d mmmm yyyy', true, 'ru'));
        $this->assertSame(1, $formatter->builds);
        $this->assertSame('2024-01-02 00:00:00', $formatter->format(45292.9999999999, $mask, true));
        $this->assertSame('2024-01-02 03:00:00', $formatter->format(45293.125, $mask, true));
        $this->assertSame(2, $formatter->builds);
        $this->assertSame('2024-01-01 12:00:00', $formatter->format(45292.5, $mask, true));
        $this->assertSame(3, $formatter->builds);
    }

    public function testSyntheticExcelDaysAreSeparateCacheKeys(): void
    {
        $formatter = new CountingCalendarFormatter();
        foreach ([[59, 'Wednesday, 28 February 1900'], [60, 'Wednesday, 29 February 1900'],
            [61, 'Thursday, 1 March 1900'], [0, 'Sunday, 0 January 1900']] as [$day, $expected]) {
            $this->assertSame($expected, $formatter->format($day, 'dddd, d mmmm yyyy', true));
            $this->assertSame($expected, $formatter->format($day + 0.5, 'dddd, d mmmm yyyy', true));
        }
        $this->assertSame(4, $formatter->builds);
        $this->assertSame('9999-12-31', $formatter->format(2958465, 'yyyy-mm-dd', true));
        $this->assertSame('10000-01-01', $formatter->format(2958465.99999999, 'yyyy-mm-dd', true));
    }

    public function testTimeAndElapsedFormatsDoNotBuildCalendarParts(): void
    {
        $formatter = new CountingCalendarFormatter();
        $this->assertSame('6:00 PM', $formatter->format(61.75, 'h:mm AM/PM', true));
        $this->assertSame('36:00:00', $formatter->format(1.5, '[h]:mm:ss', true));
        $this->assertSame('18:00', $formatter->format(61.75, 'hh:mm', true));
        $this->assertSame(0, $formatter->builds);
        $this->assertNull($formatter->epoch());
        $this->assertSame('-1', $formatter->format(-1, '[h]:mm', true));
        $this->assertSame('2958466', $formatter->format(2958466, 'hh:mm', true));
    }

    public function testCalendarCachesAreIndependent(): void
    {
        $first = new CountingCalendarFormatter();
        $second = new CountingCalendarFormatter();
        $this->assertSame('2024-01-01 00:00', $first->format(45292, 'yyyy-mm-dd hh:mm', true));
        $this->assertSame('2024-01-02 00:00', $second->format(45293, 'yyyy-mm-dd hh:mm', true));
        $this->assertSame('2024-01-01 12:00', $first->format(45292.5, 'yyyy-mm-dd hh:mm', true));
        $this->assertSame(1, $first->builds);
        $this->assertSame(1, $second->builds);
    }
}
