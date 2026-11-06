<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Style\DisplayFormatter;
use PHPUnit\Framework\TestCase;

class UtcCalendarFormatter extends DisplayFormatter
{
    public function parts(int $day): array { return $this->buildCalendarParts($day); }
    public function fallback(int $day): DateTimeImmutable { return $this->calendarDateTime($day); }
    public function epoch(): ?DateTimeImmutable { return $this->dateEpoch; }
}

final class AutoWidthUtcCalendarTest extends TestCase
{
    public function testCalendarMatchesTheDateTimeFallbackAcrossTheSupportedRange(): void
    {
        $formatter = new UtcCalendarFormatter();
        $days = [0, 1, 58, 59, 60, 61, 365, 366, 367, 25568, 25569, 25570, 45292, 2958465, 2958466];
        // Cover leap-year and century boundaries without assuming a timestamp range.
        foreach (['2000-02-28', '2000-02-29', '2000-03-01', '2100-02-28', '2100-03-01', '2400-02-29'] as $date) {
            $days[] = (int)(new DateTimeImmutable('1899-12-30', new DateTimeZone('UTC')))
                ->diff(new DateTimeImmutable($date, new DateTimeZone('UTC')))->days;
        }
        for ($day = 0; $day <= 2958465; $day += 7919) { $days[] = $day; }
        foreach ($days as $day) {
            $parts = explode(' ', $formatter->fallback($day)->format('Y n j w'));
            $expected = [$day === 0 ? 1900 : (int)$parts[0], $day === 0 ? 1 : ($day === 60 ? 2 : (int)$parts[1]),
                $day === 0 ? 0 : ($day === 60 ? 29 : (int)$parts[2]), (int)$parts[3]];
            $this->assertSame($expected, $formatter->parts($day), 'Excel day ' . $day);
        }
    }

    public function testUtcConversionDoesNotAllocateTheFallbackEpochOn64BitPhp(): void
    {
        $formatter = new UtcCalendarFormatter();
        $timezone = date_default_timezone_get();
        try {
            foreach (['Pacific/Honolulu', 'Europe/Moscow', 'UTC'] as $name) {
                date_default_timezone_set($name);
                $this->assertSame([2024, 1, 1, 1], $formatter->parts(45292));
                $this->assertSame([1900, 2, 29, 3], $formatter->parts(60));
                $this->assertSame([1900, 1, 0, 0], $formatter->parts(0));
                $this->assertSame([10000, 1, 1, 6], $formatter->parts(2958466));
            }
            if (PHP_INT_SIZE >= 8) { $this->assertNull($formatter->epoch()); }
        }
        finally { date_default_timezone_set($timezone); }
    }
}
