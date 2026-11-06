<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Style\DisplayFormatter;
use PHPUnit\Framework\TestCase;

class CountingFormatPlanFormatter extends DisplayFormatter
{
    public int $numberPlans = 0;
    public int $datePlans = 0;
    public bool $fail = false;
    public bool $nested = false;
    public string $innerResult = '';

    protected function prepareNumber(array $tokens): array
    {
        ++$this->numberPlans;
        return parent::prepareNumber($tokens);
    }

    protected function prepareDate(array $tokens): array
    {
        ++$this->datePlans;
        return parent::prepareDate($tokens);
    }

    protected function number(float $number, array $tokens, string $locale): ?string
    {
        if ($this->fail) {
            throw new RuntimeException('Rendering failed');
        }
        if ($this->nested) {
            $this->nested = false;
            $this->innerResult = $this->format(61, 'yyyy-mm-dd', true, 'en');
        }
        return parent::number($number, $tokens, $locale);
    }

    public function hasActivePlan(): bool
    {
        return $this->activeSection !== null;
    }
}

final class AutoWidthFormatPlanTest extends TestCase
{
    public function testPlansArePreparedOnceAndDoNotCaptureTheLocale(): void
    {
        $formatter = new CountingFormatPlanFormatter();
        for ($i = 0; $i < 20; ++$i) {
            $this->assertSame(number_format($i / 10, 3, '.', ','), $formatter->format($i / 10, '#,##0.000', true, 'en'));
        }
        $this->assertSame("1\xc2\xa0234,500", $formatter->format(1234.5, '#,##0.000', true, 'ru'));
        $this->assertSame(1, $formatter->numberPlans);
        $this->assertSame('1900-03-01 12:00:00', $formatter->format(61.5, 'yyyy-mm-dd hh:mm:ss', true));
        $this->assertSame('1900-03-02 12:00:00', $formatter->format(62.5, 'yyyy-mm-dd hh:mm:ss', true));
        $this->assertSame('0000-00-00 00:00:00', $formatter->dateWidthSample('yyyy-mm-dd hh:mm:ss'));
        $this->assertSame(1, $formatter->datePlans);
    }

    public function planValues(): array
    {
        return [
            [-12.5, '0.0;[Red](0.0);"zero"', '(12.5)'],
            [-12.5, '0.0;[<0](0.0);"zero"', '(-12.5)'],
            [5, '0;[>10]0.00;"other"', '5'],
            [15, '0;[>10]0.00;"other"', '15.00'],
            ['000', '[<0]0;[>0]0', '000'],
            [0, '0.0;0.0;', ''],
            [1234567, '0.0,,', '1.2'],
            [1234567, '0.00e-000', '1.23e006'],
            [-0.00001, '0.000', '-0.000'],
            [1.2, '0.??', '1.2 '],
            [0, '#.##', ''],
            [61.75, 'h:mm "AM/PM"', '6:00 AM/PM'],
            [61.75, 'h:mm A/P', '6:00 P'],
            [61.75, 'mm hh mm ss', '03 18 00 00'],
            [61.75, 'mmmm "x" mm:ss', 'March x 00:00'],
            [61.75, 'yyyy-mm-dd hh:mm:ss', '1900-03-01 18:00:00'],
            ['abc', '0;0;0;"text: "@;yyyy', 'text: abc'],
        ];
    }

    /** @dataProvider planValues */
    public function testPreparedPlansPreserveExistingFormatting($value, string $format, string $expected): void
    {
        $formatter = new DisplayFormatter();
        $this->assertSame($expected, $formatter->format($value, $format, true, 'en'));
    }

    public function testNestedRenderingAndExceptionsRestoreTheActivePlan(): void
    {
        $formatter = new CountingFormatPlanFormatter();
        $formatter->nested = true;
        $this->assertSame('12.500', $formatter->format(12.5, '0.000', true));
        $this->assertSame('1900-03-01', $formatter->innerResult);
        $this->assertSame(1, $formatter->numberPlans);
        $this->assertFalse($formatter->hasActivePlan());
        $formatter->fail = true;
        try {
            $formatter->format(12.5, '0.000', true);
            $this->fail('The custom renderer must throw');
        }
        catch (RuntimeException $error) {
            $this->assertSame('Rendering failed', $error->getMessage());
        }
        $this->assertFalse($formatter->hasActivePlan());
        $formatter->fail = false;
        $this->assertSame('12.500', $formatter->format(12.5, '0.000', true));
        $this->assertSame(1, $formatter->numberPlans);
    }

    public function testLargePlansAreNotRetained(): void
    {
        $formatter = new CountingFormatPlanFormatter();
        $format = str_repeat('0', 257);
        $this->assertSame('12', $formatter->format(12, $format, true));
        $this->assertSame('12', $formatter->format(12, $format, true));
        $this->assertSame(2, $formatter->numberPlans);
    }
}
