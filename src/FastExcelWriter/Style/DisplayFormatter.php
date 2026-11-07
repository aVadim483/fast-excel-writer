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
    protected ?\DateTimeImmutable $dateEpoch = null;
    protected ?array $activeSection = null;
    protected ?int $lastCalendarDay = null;
    protected array $lastCalendarParts = [];

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
        $plan = $this->formatPlan($format);
        $sections = $plan['sections'];
        $sample = '';
        $hasDate = false;
        $supported = $plan['count'] === 1 && !$plan['conditional'];
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

    /** Boolean captions follow the workbook language, independently of number masks. */
    public static function booleanLabels(string $locale): array
    {
        static $labels = [
            'cs' => ['NEPRAVDA', 'PRAVDA'], 'da' => ['FALSK', 'SAND'],
            'de' => ['FALSCH', 'WAHR'], 'es' => ['FALSO', 'VERDADERO'],
            'fi' => ['EPÄTOSI', 'TOSI'], 'fr' => ['FAUX', 'VRAI'],
            'hu' => ['HAMIS', 'IGAZ'], 'it' => ['FALSO', 'VERO'],
            'nl' => ['ONWAAR', 'WAAR'], 'no' => ['USANN', 'SANN'],
            'pl' => ['FAŁSZ', 'PRAWDA'], 'pt' => ['FALSO', 'VERDADEIRO'],
            'ru' => ['ЛОЖЬ', 'ИСТИНА'], 'sv' => ['FALSKT', 'SANT'],
            'tr' => ['YANLIŞ', 'DOĞRU'],
        ];
        return $labels[strtolower(substr($locale, 0, 2))] ?? ['FALSE', 'TRUE'];
    }

    public function format($value, string $format, bool $numeric = false, string $locale = 'en'): string
    {
        if (is_bool($value)) {
            return self::booleanLabels($locale)[(int)$value];
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
        $plan = $this->formatPlan($format);
        $sections = $plan['sections'];
        if (!$numeric || !is_numeric($value)) {
            $section = $plan['text'] === null ? null : $sections[$plan['text']];
            if ($section === null || ($plan['count'] < 4 && !$section['has_text'])) {
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
        if ($plan['conditional']) {
            for ($i = 0; $i < min(3, $plan['count']); ++$i) {
                if ($sections[$i]['condition'] !== null && $this->matches($number, $sections[$i]['condition'])) {
                    $section = $sections[$i];
                    break;
                }
            }
            if ($section === null && $plan['fallback'] !== null) {
                $section = $sections[$plan['fallback']];
            }
            if ($section === null) {
                return $text;
            }
        }
        else {
            $index = $number < 0 && $plan['count'] > 1 ? 1 : ($number == 0 && $plan['count'] > 2 ? 2 : 0);
            $section = $sections[$index];
            $absolute = $index === 1;
        }
        if ($section['locale'] !== '') {
            $locale = $section['locale'];
        }
        if ($section['kind'] === 'general') {
            return str_replace('.', $this->separators($locale)[0], $text);
        }
        $previous = $this->activeSection;
        $this->activeSection = $section;
        try {
            if ($section['kind'] === 'date') {
                return $this->formatDate($number, $section['tokens'], $locale, $format) ?? $text;
            }
            return $this->number($absolute ? abs($number) : $number, $section['tokens'], $locale) ?? $text;
        }
        finally {
            $this->activeSection = $previous;
        }
    }

    /** Keep compiled sections in the existing bounded format cache. */
    protected function formatPlan(string $format): array
    {
        $plan = $this->formats->get($format);
        if ($plan !== null) {
            return $plan;
        }
        $parsed = $this->parse($format);
        $count = count($parsed);
        $plan = ['count' => $count, 'conditional' => false, 'fallback' => null,
            'text' => $count >= 4 ? 3 : ($count === 1 ? 0 : null), 'sections' => []];
        $tokenCount = 0;
        foreach ($parsed as $index => $section) {
            if ($index >= 4) {
                break;
            }
            $tokenCount += count($section['tokens']);
            if ($index < 3) {
                if ($section['condition'] !== null) {
                    $plan['conditional'] = true;
                }
                elseif ($plan['fallback'] === null) {
                    $plan['fallback'] = $index;
                }
            }
            $section['has_text'] = strpos($section['code'], '@') !== false;
            $section['kind'] = strcasecmp($section['code'], 'general') === 0 ? 'general' : 'number';
            if ($section['kind'] !== 'general') {
                foreach ($section['tokens'] as $token) {
                    if ($token[0] === 'date' || $token[0] === 'elapsed') {
                        $section['kind'] = 'date';
                        break;
                    }
                }
            }
            if ($section['kind'] === 'date') {
                $section['date_plan'] = $this->prepareDate($section['tokens']);
            }
            elseif ($section['kind'] === 'number') {
                $section['number_plan'] = $this->prepareNumber($section['tokens']);
            }
            $plan['sections'][$index] = $section;
        }
        if ($tokenCount <= 256) {
            $this->formats->put($format, $plan);
        }
        return $plan;
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

    protected function prepareNumber(array $tokens): array
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
            return ['kind' => 'literal', 'text' => $this->literal($tokens)];
        }
        while (isset($tokens[$last + 1]) && $tokens[$last + 1] === ['code', ',']) {
            ++$last;
        }
        $pattern = '';
        for ($i = $first; $i <= $last; ++$i) {
            if ($tokens[$i][0] !== 'code') {
                return ['kind' => 'unsupported'];
            }
            $pattern .= $tokens[$i][1];
        }
        if (!preg_match('/^([#0?,]+)(?:\.([#0?]+))?(?:([Ee])([+-])(0+))?(,*)$/', $pattern, $parts)) {
            return ['kind' => 'unsupported'];
        }
        $integerMask = $parts[1];
        $scale = strlen($integerMask) - strlen(rtrim($integerMask, ','));
        $integerMask = rtrim($integerMask, ',');
        $scale += strlen($parts[6] ?? '');
        $fractionMask = $parts[2] ?? '';
        $decimals = strlen($fractionMask);
        if ($decimals > 30 || strlen($integerMask) > 100) {
            return ['kind' => 'unsupported'];
        }
        return ['kind' => 'number', 'multiplier' => pow(100, $percent) / pow(1000, $scale),
            'fraction' => $fractionMask, 'decimals' => $decimals, 'zeroes' => substr_count($integerMask, '0'),
            'trim_limit' => $decimals - strspn(strrev($fractionMask), '#?'), 'scientific_format' => '%.' . $decimals . 'E',
            'grouped' => strpos($integerMask, ',') !== false, 'prefix' => $this->literal(array_slice($tokens, 0, $first)),
            'suffix' => $this->literal(array_slice($tokens, $last + 1)), 'exponent' => $parts[3] ?? '',
            'exponent_sign' => $parts[4] ?? '', 'exponent_zeroes' => strlen($parts[5] ?? '')];
    }

    protected function number(float $number, array $tokens, string $locale): ?string
    {
        $section = $this->activeSection;
        $plan = $section !== null && isset($section['number_plan']) && $tokens === $section['tokens']
            ? $section['number_plan'] : $this->prepareNumber($tokens);
        if ($plan['kind'] === 'literal') {
            return $plan['text'];
        }
        if ($plan['kind'] === 'unsupported') {
            return null;
        }
        $number *= $plan['multiplier'];
        if (!is_finite($number)) {
            return null;
        }
        $fractionMask = $plan['fraction'];
        $decimals = $plan['decimals'];
        $exponent = '';
        if ($plan['exponent'] !== '') {
            $scientific = explode('E', sprintf($plan['scientific_format'], abs($number)));
            $number = ($number < 0 ? -1 : 1) * (float)str_replace(',', '.', $scientific[0]);
            $power = (int)$scientific[1];
            $exponent = $plan['exponent'] . ($power < 0 ? '-' : ($plan['exponent_sign'] === '+' ? '+' : ''))
                . str_pad((string)abs($power), $plan['exponent_zeroes'], '0', STR_PAD_LEFT);
        }
        [$decimal, $thousands] = $this->separators($locale);
        $formatted = explode('.', number_format(abs($number), $decimals, '.', ''));
        $integer = str_pad($formatted[0], $plan['zeroes'], '0', STR_PAD_LEFT);
        if ($integer === '0' && $plan['zeroes'] === 0) {
            $integer = '';
        }
        if ($plan['grouped']) {
            $integer = preg_replace('/\B(?=(\d{3})+(?!\d))/', $thousands, $integer);
        }
        $fraction = $formatted[1] ?? '';
        for ($i = $decimals - 1; $i >= $plan['trim_limit']; --$i) {
            if (($fraction[$i] ?? '') !== '0') {
                break;
            }
            $fraction = substr($fraction, 0, $i) . ($fractionMask[$i] === '?' ? ' ' : '');
        }
        $output = ($number < 0 ? '-' : '') . $integer . ($fraction !== '' ? $decimal . $fraction : '') . $exponent;
        return $plan['prefix'] . $output . $plan['suffix'];
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

    protected function prepareDate(array $tokens): array
    {
        $parts = [];
        $previous = [];
        $next = [];
        $last = '';
        $amPm = false;
        foreach ($tokens as $index => $token) {
            $parts[$index] = strtolower($token[1]);
            $previous[$index] = $last;
            if ($parts[$index] === 'am/pm' || $parts[$index] === 'a/p') {
                $amPm = true;
            }
            if ($token[0] === 'date' || $token[0] === 'elapsed') {
                $last = $parts[$index];
            }
        }
        $last = '';
        for ($index = count($tokens) - 1; $index >= 0; --$index) {
            $next[$index] = $last;
            if ($tokens[$index][0] === 'date' || $tokens[$index][0] === 'elapsed') {
                $last = $parts[$index];
            }
        }
        $plan = ['operations' => [], 'am_pm' => $amPm, 'calendar' => false];
        foreach ($tokens as $index => $token) {
            $part = $parts[$index];
            $length = strlen($part);
            if ($token[0] === 'elapsed') {
                $plan['operations'][] = ['elapsed', $length, $part[0] === 'h' ? 3600 : ($part[0] === 'm' ? 60 : 1)];
                continue;
            }
            if ($token[0] !== 'date') {
                $lastIndex = count($plan['operations']) - 1;
                if ($lastIndex >= 0 && $plan['operations'][$lastIndex][0] === 'literal') {
                    $plan['operations'][$lastIndex][1] .= $token[1];
                }
                else {
                    $plan['operations'][] = ['literal', $token[1]];
                }
                continue;
            }
            if ($part === 'am/pm' || $part === 'a/p') {
                $plan['operations'][] = ['am_pm', $part === 'a/p'];
                continue;
            }
            if ($part[0] === 'y') {
                $kind = 'year';
            }
            elseif ($part[0] === 'd') {
                $kind = $length >= 3 ? 'weekday' : 'day';
            }
            elseif ($part[0] === 'm') {
                $kind = ($previous[$index][0] ?? '') === 'h' || ($next[$index][0] ?? '') === 's' ? 'minute' : 'month';
            }
            else {
                $kind = $part[0] === 'h' ? 'hour' : 'second';
            }
            if (in_array($kind, ['year', 'month', 'day', 'weekday'], true)) {
                $plan['calendar'] = true;
            }
            $plan['operations'][] = [$kind, $length];
        }
        $plan['native_mask'] = $plan['calendar'] ? $this->nativeDateMask($plan) : null;
        return $plan;
    }

    /** Compile only calendar operations with an exact PHP date-format equivalent. */
    protected function nativeDateMask(array $plan): ?string
    {
        $mask = '';
        foreach ($plan['operations'] as $operation) {
            [$kind, $length] = $operation;
            switch ($kind) {
                case 'literal':
                    if (strpos($operation[1], "\0") !== false) {
                        return null;
                    }
                    // Escape every byte, including date-format letters and backslashes.
                    for ($index = 0, $literalLength = strlen($operation[1]); $index < $literalLength; ++$index) {
                        $mask .= '\\' . $operation[1][$index];
                    }
                    break;
                case 'year':
                    $mask .= $length <= 2 ? 'y' : 'Y';
                    break;
                case 'day':
                    $mask .= $length === 1 ? 'j' : 'd';
                    break;
                case 'weekday':
                    $mask .= $length === 3 ? 'D' : 'l';
                    break;
                case 'month':
                    if ($length > 4) {
                        return null;
                    }
                    $mask .= [1 => 'n', 2 => 'm', 3 => 'M', 4 => 'F'][$length];
                    break;
                case 'minute':
                case 'second':
                    if ($length !== 2) {
                        return null;
                    }
                    $mask .= $kind === 'minute' ? 'i' : 's';
                    break;
                case 'hour':
                    if ($length > 2) {
                        return null;
                    }
                    $mask .= $plan['am_pm'] ? ($length === 1 ? 'g' : 'h') : ($length === 1 ? 'G' : 'H');
                    break;
                case 'am_pm':
                    if ($operation[1]) {
                        return null;
                    }
                    $mask .= 'A';
                    break;
                default:
                    return null;
            }
        }
        return $mask;
    }

    /** Cache only the effective Excel day, including midnight carry and the fake leap day. */
    protected function calendarParts(int $day): array
    {
        if ($day !== $this->lastCalendarDay) {
            $this->lastCalendarParts = $this->buildCalendarParts($day);
            $this->lastCalendarDay = $day;
        }
        return $this->lastCalendarParts;
    }

    protected function buildCalendarParts(int $day): array
    {
        if (PHP_INT_SIZE >= 8) {
            // Before serial 60, Excel's leap-day correction has not taken effect yet.
            $parts = explode(' ', gmdate('Y n j w', ($day - 25569 + ($day < 60 ? 1 : 0)) * 86400));
        }
        else {
            $parts = explode(' ', $this->calendarDateTime($day)->format('Y n j w'));
        }
        return [$day === 0 ? 1900 : (int)$parts[0], $day === 0 ? 1 : ($day === 60 ? 2 : (int)$parts[1]),
            $day === 0 ? 0 : ($day === 60 ? 29 : (int)$parts[2]), (int)$parts[3]];
    }

    /** Keep the full calendar range on platforms with 32-bit timestamps. */
    protected function calendarDateTime(int $day): \DateTimeImmutable
    {
        if ($this->dateEpoch === null) {
            $this->dateEpoch = new \DateTimeImmutable('1899-12-31', new \DateTimeZone('UTC'));
        }
        return $this->dateEpoch->modify('+' . ($day - ($day >= 60 ? 1 : 0)) . ' days');
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
        $section = $this->activeSection;
        $plan = $section !== null && isset($section['date_plan']) && $tokens === $section['tokens']
            ? $section['date_plan'] : $this->prepareDate($tokens);
        if (PHP_INT_SIZE >= 8 && $locale === 'en' && isset($plan['native_mask']) && $day !== 0 && $day !== 60) {
            return gmdate($plan['native_mask'], ($day - 25569 + ($day < 60 ? 1 : 0)) * 86400 + $seconds);
        }
        $calendar = $plan['calendar'] ? $this->calendarParts($day) : null;
        $hour = intdiv($seconds, 3600);
        $minute = intdiv($seconds % 3600, 60);
        $second = $seconds % 60;
        $amPm = $plan['am_pm'];
        $months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
        $days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        if (substr(strtolower($locale), 0, 2) === 'ru') {
            $months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
            $days = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
        }
        $output = '';
        foreach ($plan['operations'] as $operation) {
            $kind = $operation[0];
            $length = $operation[1];
            switch ($kind) {
                case 'elapsed':
                    $output .= str_pad((string)(int)floor(($day * 86400 + $seconds) / $operation[2]), $length, '0', STR_PAD_LEFT);
                    break;
                case 'literal':
                    $output .= $operation[1];
                    break;
                case 'am_pm':
                    $output .= $operation[1] ? ($hour < 12 ? 'A' : 'P') : ($hour < 12 ? 'AM' : 'PM');
                    break;
                case 'year':
                    $output .= $length <= 2 ? str_pad((string)($calendar[0] % 100), 2, '0', STR_PAD_LEFT) : $calendar[0];
                    break;
                case 'day':
                    $output .= str_pad((string)$calendar[2], $length, '0', STR_PAD_LEFT);
                    break;
                case 'weekday':
                    $output .= $length === 3 ? mb_substr($days[$calendar[3]], 0, 3) : $days[$calendar[3]];
                    break;
                case 'month':
                    $output .= $length >= 3 ? ($length === 3 ? mb_substr($months[$calendar[1] - 1], 0, 3) : ($length === 5 ? mb_substr($months[$calendar[1] - 1], 0, 1) : $months[$calendar[1] - 1]))
                        : str_pad((string)$calendar[1], $length, '0', STR_PAD_LEFT);
                    break;
                case 'minute':
                    $output .= str_pad((string)$minute, $length, '0', STR_PAD_LEFT);
                    break;
                case 'hour':
                    $output .= str_pad((string)($amPm ? (($hour % 12) ?: 12) : $hour), $length, '0', STR_PAD_LEFT);
                    break;
                case 'second':
                    $output .= str_pad((string)$second, $length, '0', STR_PAD_LEFT);
                    break;
            }
        }
        return $output;
    }
}
