<?php

namespace avadim\FastExcelWriter\Style;

/** Display-only subset of Excel formats. Unsupported masks return the original text. */
class DisplayFormatter
{
    protected WidthCache $formats;
    protected WidthCache $dateSamples;
    protected ?float $lastDateValue = null;
    protected string $lastDateFormat = '';
    protected string $lastDateLocale = '';
    protected ?string $lastDateText = null;

    public function __construct()
    {
        $this->formats = new WidthCache(256, 262144, 4096);
        $this->dateSamples = new WidthCache(256, 262144, 4096);
    }

    /** A numeric date mask has a reusable width sample; variable formats use the full formatter. */
    public function dateWidthSample(string $format): ?string
    {
        $cached = $this->dateSamples->get($format);
        if ($cached !== null) {
            return $cached === false ? null : $cached;
        }
        $sections = $this->formats->get($format);
        if ($sections === null) {
            $sections = $this->parse($format);
            $this->formats->put($format, $sections);
        }
        $sample = '';
        $hasDate = false;
        $supported = count($sections) === 1 && $sections[0]['condition'] === null;
        if ($supported) {
            foreach ($sections[0]['tokens'] as $token) {
                if ($token[0] === 'date') {
                    $part = strtolower($token[1]);
                    if (!preg_match('/^(?:y{1,4}|m{1,2}|d{1,2}|h{1,2}|s{1,2})$/', $part)) {
                        $supported = false;
                        break;
                    }
                    $hasDate = true;
                    $sample .= str_repeat('0', $part[0] === 'y' ? (strlen($part) <= 2 ? 2 : 4) : 2);
                }
                elseif ($token[0] === 'literal' || ($token[0] === 'code' && preg_match('/^[\s:\/.,-]$/u', $token[1]))) {
                    $sample .= $token[1];
                }
                else {
                    $supported = false;
                    break;
                }
            }
        }
        $result = $supported && $hasDate ? $sample : false;
        $this->dateSamples->put($format, $result);
        return $result === false ? null : $result;
    }

    public function format($value, string $format, bool $numeric = false, string $locale = 'en'): string
    {
        if (is_bool($value)) {
            return $value ? 'TRUE' : 'FALSE';
        }
        if ($value === null) {
            return '';
        }
        $text = (string)$value;
        if ($format === '@') {
            return $text;
        }
        if ($format === '' || strcasecmp($format, 'general') === 0) {
            return $numeric ? str_replace('.', $this->separators($locale)[0], $text) : $text;
        }
        if ($numeric && ($format === '0.00' || $format === '0.00%') && (is_int($value) || is_float($value))) {
            $number = (float)$value * ($format === '0.00%' ? 100 : 1);
            if (!is_finite($number)) {
                return $text;
            }
            // Keep the sign even when a negative value rounds to zero.
            return ($number < 0 ? '-' : '') . number_format(abs($number), 2, $this->separators($locale)[0], '')
                . ($format === '0.00%' ? '%' : '');
        }
        $sections = $this->formats->get($format);
        if ($sections === null) {
            $sections = $this->parse($format);
            $this->formats->put($format, $sections);
        }
        if (!$numeric || !is_numeric($value)) {
            $section = $sections[3] ?? (count($sections) === 1 ? $sections[0] : null);
            if ($section === null || (count($sections) < 4 && strpos($section['code'], '@') === false)) {
                return $text;
            }
            return $this->literal($section['tokens'], $text);
        }
        $number = (float)$value;
        if (!is_finite($number)) {
            return $text;
        }
        $section = null;
        $absolute = false;
        $conditional = false;
        foreach (array_slice($sections, 0, 3) as $candidate) {
            if ($candidate['condition'] !== null) {
                $conditional = true;
                if ($this->matches($number, $candidate['condition'])) {
                    $section = $candidate;
                    break;
                }
            }
        }
        if ($conditional && $section === null) {
            foreach (array_slice($sections, 0, 3) as $candidate) {
                if ($candidate['condition'] === null) {
                    $section = $candidate;
                    break;
                }
            }
            if ($section === null) {
                return $text;
            }
        }
        if (!$conditional) {
            $index = $number < 0 && count($sections) > 1 ? 1 : ($number == 0 && count($sections) > 2 ? 2 : 0);
            $section = $sections[$index];
            $absolute = $index === 1;
        }
        if ($section['locale'] !== '') {
            $locale = $section['locale'];
        }
        if (strcasecmp($section['code'], 'general') === 0) {
            return str_replace('.', $this->separators($locale)[0], $text);
        }
        foreach ($section['tokens'] as $token) {
            if ($token[0] === 'date' || $token[0] === 'elapsed') {
                return $this->formatDate($number, $section['tokens'], $locale, $format) ?? $text;
            }
        }
        return $this->number($absolute ? abs($number) : $number, $section['tokens'], $locale) ?? $text;
    }

    protected function parse(string $format): array
    {
        preg_match_all('/"[^"]*"|\\\\.|\[[^\]]*\]|_.|\*.|AM\/PM|A\/P|y+|m+|d+|h+|s+|[^;]|;/iu', $format, $matches);
        $sections = [];
        $section = ['tokens' => [], 'condition' => null, 'locale' => '', 'code' => ''];
        foreach ($matches[0] as $part) {
            if ($part === ';') {
                $sections[] = $section;
                $section = ['tokens' => [], 'condition' => null, 'locale' => '', 'code' => ''];
                continue;
            }
            $kind = 'code';
            $value = $part;
            if ($part[0] === '"') {
                $kind = 'literal';
                $value = substr($part, 1, -1);
            }
            elseif ($part[0] === '\\') {
                $kind = 'literal';
                $value = substr($part, 1);
            }
            elseif ($part[0] === '_') {
                $kind = 'literal';
                $value = ' ';
            }
            elseif ($part[0] === '*') {
                continue; // Repeated fill depends on the final column width.
            }
            elseif ($part[0] === '[') {
                if (preg_match('/^\[(<=|>=|<>|=|<|>)(-?[0-9.]+)\]$/', $part, $condition)) {
                    $section['condition'] = [$condition[1], (float)$condition[2]];
                    continue;
                }
                if (preg_match('/^\[([hms]+)\]$/i', $part, $elapsed)) {
                    $kind = 'elapsed';
                    $value = strtolower($elapsed[1]);
                }
                elseif (preg_match('/^\[\$([^\]-]*)(?:-([0-9a-f]+))?\]$/i', $part, $currency)) {
                    $kind = 'literal';
                    $value = $currency[1];
                    $locales = ['409' => 'en', '407' => 'de', '419' => 'ru', '40c' => 'fr'];
                    $section['locale'] = $locales[strtolower($currency[2] ?? '')] ?? '';
                }
                else {
                    continue; // Colors and unsupported metadata do not add visible text.
                }
            }
            elseif (preg_match('/^(?:[ymdhs]+|AM\/PM|A\/P)$/i', $part)) {
                $kind = 'date';
            }
            $section['tokens'][] = [$kind, $value];
            if ($kind === 'code') {
                $section['code'] .= $value;
            }
        }
        $sections[] = $section;
        return $sections;
    }

    protected function matches(float $number, array $condition): bool
    {
        switch ($condition[0]) {
            case '<': return $number < $condition[1];
            case '>': return $number > $condition[1];
            case '<=': return $number <= $condition[1];
            case '>=': return $number >= $condition[1];
            case '<>': return $number != $condition[1];
            default: return $number == $condition[1];
        }
    }

    protected function literal(array $tokens, string $text = ''): string
    {
        $result = '';
        foreach ($tokens as $token) {
            $result .= $token[0] === 'code' ? str_replace('@', $text, $token[1]) : $token[1];
        }
        return $result;
    }

    protected function separators(string $locale): array
    {
        $language = strtolower(substr($locale, 0, 2));
        if ($language === 'de') {
            return [',', '.'];
        }
        if ($language === 'ru' || $language === 'fr') {
            return [',', "\xc2\xa0"];
        }
        return ['.', ','];
    }

    protected function number(float $number, array $tokens, string $locale): ?string
    {
        $first = null;
        $last = null;
        $percent = 0;
        foreach ($tokens as $i => $token) {
            if ($token[0] === 'code' && strpbrk($token[1], '0#?') !== false) {
                $first = $first ?? $i;
                $last = $i;
            }
            if ($token === ['code', '%']) {
                ++$percent;
            }
        }
        if ($first === null) {
            return $this->literal($tokens);
        }
        while (isset($tokens[$last + 1]) && $tokens[$last + 1] === ['code', ',']) {
            ++$last;
        }
        $pattern = '';
        for ($i = $first; $i <= $last; ++$i) {
            if ($tokens[$i][0] !== 'code') {
                return null;
            }
            $pattern .= $tokens[$i][1];
        }
        if (!preg_match('/^([#0?,]+)(?:\.([#0?]+))?(?:([Ee])([+-])(0+))?(,*)$/', $pattern, $parts)) {
            return null;
        }
        $integerMask = $parts[1];
        $scale = strlen($integerMask) - strlen(rtrim($integerMask, ','));
        $integerMask = rtrim($integerMask, ',');
        $scale += strlen($parts[6] ?? '');
        $number *= pow(100, $percent) / pow(1000, $scale);
        if (!is_finite($number)) {
            return null;
        }
        $fractionMask = $parts[2] ?? '';
        $decimals = strlen($fractionMask);
        if ($decimals > 30 || strlen($integerMask) > 100) {
            return null;
        }
        $exponent = '';
        if (!empty($parts[3])) {
            $scientific = explode('E', sprintf('%.' . $decimals . 'E', abs($number)));
            $number = ($number < 0 ? -1 : 1) * (float)str_replace(',', '.', $scientific[0]);
            $power = (int)$scientific[1];
            $exponent = $parts[3] . ($power < 0 ? '-' : ($parts[4] === '+' ? '+' : ''))
                . str_pad((string)abs($power), strlen($parts[5]), '0', STR_PAD_LEFT);
        }
        [$decimal, $thousands] = $this->separators($locale);
        $formatted = explode('.', number_format(abs($number), $decimals, '.', ''));
        $integer = str_pad($formatted[0], substr_count($integerMask, '0'), '0', STR_PAD_LEFT);
        if ($integer === '0' && strpos($integerMask, '0') === false) {
            $integer = '';
        }
        if (strpos($integerMask, ',') !== false) {
            $integer = preg_replace('/\B(?=(\d{3})+(?!\d))/', $thousands, $integer);
        }
        $fraction = $formatted[1] ?? '';
        for ($i = $decimals - 1; $i >= 0 && $i < strlen($fractionMask); --$i) {
            if ($fractionMask[$i] === '0' || ($fraction[$i] ?? '') !== '0') {
                break;
            }
            $fraction = substr($fraction, 0, $i) . ($fractionMask[$i] === '?' ? ' ' : '');
        }
        $output = ($number < 0 ? '-' : '') . $integer . ($fraction !== '' ? $decimal . $fraction : '') . $exponent;
        return $this->literal(array_slice($tokens, 0, $first)) . $output
            . $this->literal(array_slice($tokens, $last + 1));
    }

    /** Cache one exact date value; time fractions and section selection remain significant. */
    protected function formatDate(float $serial, array $tokens, string $locale, string $format): ?string
    {
        if ($serial === $this->lastDateValue && $format === $this->lastDateFormat && $locale === $this->lastDateLocale) {
            return $this->lastDateText;
        }
        $text = $this->date($serial, $tokens, $locale);
        if ($text !== null && strlen($format) <= 4096 && strlen($locale) <= 256 && strlen($text) <= 4096) {
            $this->lastDateFormat = $format;
            $this->lastDateLocale = $locale;
            $this->lastDateText = $text;
            $this->lastDateValue = $serial;
        }
        return $text;
    }

    protected function date(float $serial, array $tokens, string $locale): ?string
    {
        if ($serial < 0 || $serial >= 2958466) {
            return null;
        }
        $day = (int)floor($serial);
        $seconds = (int)round(($serial - $day) * 86400);
        if ($seconds >= 86400) {
            ++$day;
            $seconds = 0;
        }
        $date = (new \DateTimeImmutable('1899-12-31', new \DateTimeZone('UTC')))
            ->modify('+' . ($day - ($day >= 60 ? 1 : 0)) . ' days');
        $year = $day === 0 ? 1900 : (int)$date->format('Y');
        $month = $day === 0 ? 1 : ($day === 60 ? 2 : (int)$date->format('n'));
        $monthDay = $day === 60 ? 29 : ($day === 0 ? 0 : (int)$date->format('j'));
        $hour = intdiv($seconds, 3600);
        $minute = intdiv($seconds % 3600, 60);
        $second = $seconds % 60;
        $amPm = false;
        foreach ($tokens as $token) {
            if (in_array(strtolower($token[1]), ['am/pm', 'a/p'], true)) {
                $amPm = true;
            }
        }
        $months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
        $days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        if (substr(strtolower($locale), 0, 2) === 'ru') {
            $months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
            $days = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
        }
        $output = '';
        foreach ($tokens as $i => $token) {
            $part = strtolower($token[1]);
            $length = strlen($part);
            if ($token[0] === 'elapsed') {
                $divisor = $part[0] === 'h' ? 3600 : ($part[0] === 'm' ? 60 : 1);
                $output .= str_pad((string)(int)floor(($day * 86400 + $seconds) / $divisor), $length, '0', STR_PAD_LEFT);
            }
            elseif ($token[0] !== 'date') {
                $output .= $token[1];
            }
            elseif ($part === 'am/pm' || $part === 'a/p') {
                $output .= $part === 'a/p' ? ($hour < 12 ? 'A' : 'P') : ($hour < 12 ? 'AM' : 'PM');
            }
            elseif ($part[0] === 'y') {
                $output .= $length <= 2 ? str_pad((string)($year % 100), 2, '0', STR_PAD_LEFT) : $year;
            }
            elseif ($part[0] === 'd') {
                $output .= $length >= 3 ? ($length === 3 ? mb_substr($days[(int)$date->format('w')], 0, 3) : $days[(int)$date->format('w')])
                    : str_pad((string)$monthDay, $length, '0', STR_PAD_LEFT);
            }
            elseif ($part[0] === 'm') {
                $previous = '';
                $next = '';
                for ($j = $i - 1; $j >= 0; --$j) {
                    if ($tokens[$j][0] === 'date' || $tokens[$j][0] === 'elapsed') { $previous = strtolower($tokens[$j][1]); break; }
                }
                for ($j = $i + 1; $j < count($tokens); ++$j) {
                    if ($tokens[$j][0] === 'date' || $tokens[$j][0] === 'elapsed') { $next = strtolower($tokens[$j][1]); break; }
                }
                if (($previous !== '' && $previous[0] === 'h') || ($next !== '' && $next[0] === 's')) {
                    $output .= str_pad((string)$minute, $length, '0', STR_PAD_LEFT);
                }
                else {
                    $output .= $length >= 3 ? ($length === 3 ? mb_substr($months[$month - 1], 0, 3) : ($length === 5 ? mb_substr($months[$month - 1], 0, 1) : $months[$month - 1]))
                        : str_pad((string)$month, $length, '0', STR_PAD_LEFT);
                }
            }
            elseif ($part[0] === 'h') {
                $output .= str_pad((string)($amPm ? (($hour % 12) ?: 12) : $hour), $length, '0', STR_PAD_LEFT);
            }
            elseif ($part[0] === 's') {
                $output .= str_pad((string)$second, $length, '0', STR_PAD_LEFT);
            }
        }
        return $output;
    }
}
