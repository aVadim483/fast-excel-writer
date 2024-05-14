<?php

namespace avadim\FastExcelWriter;

class Font
{
    // constants for auo width
    protected const WIDTH_LOWER_CASE_LETTER = 1.1;
    protected const WIDTH_UPPER_CASE_LETTER = 1.31;
    protected const WIDTH_WIDE_LETTER = 1.75;
    protected const WIDTH_DOTS_SYMBOLS = 0.50;
    protected const WIDTH_PADDING = 0.81;

    protected static array $fontWidths = [
        'Arial' => [
            'dots' => 0.50,
            'lower' => 1.1,
            'upper' => 1.31,
            'wide' => 1.75,
        ],
        'Calibri' => [
            'dots' => 0.60,
            'lower' => 1.07,
            'upper' => 1.21,
            'wide' => 1.71,
        ],
    ];

    /**
     * @param string $fontName
     * @param int|float $fontSize
     * @param string $value
     * @param string|null $numberFormat
     *
     * @return float
     */
    public static function calcTextWidth(string $fontName, $fontSize, string $value, ?string $numberFormat = null): float
    {
        static $cache = [];

        $key = '[[[' . $fontName . ':' . $fontSize . ']]][[[' . $numberFormat . ']]][[[' . $value . ']]]';
        if (isset($cache[$key])) {
            $len = $cache[$key];
        }
        else {
            $len = self::_calcWidth($value, $fontName, $fontSize);
            if ($numberFormat !== 'GENERAL' && $numberFormat !== '0' && $numberFormat != '@') {
                $numberFormat = Excel::_formatValue($value, $numberFormat);
                $len = max($len, self::_calcWidth(str_replace('\\', '', $numberFormat), $fontName, $fontSize, true));
            }
            $cache[$key] = $len;
        }

        return $len;
    }

    /**
     * @param string $str
     * @param string $fontName
     * @param int|float $fontSize
     * @param bool|null $numFormat
     *
     * @return float
     */
    protected static function _calcWidth(string $str, string $fontName, $fontSize, ?bool $numFormat = false): float
    {
        if ($numFormat && strpos($str, ';')) {
            $lenArray = [];
            foreach (explode(';', $str) as $part) {
                $lenArray[] = self::_calcWidth($part, $fontName, $fontSize);
            }

            return max(...$lenArray);
        }

        $len = mb_strlen($str);
        $wideCount = 0;
        $upperCount = 0;
        $dotsCount = 0;
        if (preg_match_all("/[,.\-:';`IiJjfl\[\]\(\)\{\}]/u", $str, $matches)) {
            $dotsCount = count($matches[0]);
            $str = preg_replace("/[,.\-:';`IiJjfl\[\]\(\)\{\}]/u", '', $str);
        }
        if (preg_match_all("/[@%&WMQ]/u", $str, $matches)) {
            $wideCount = count($matches[0]);
            $str = preg_replace("/[@%&WMQ]/u", '', $str);
        }
        if (preg_match_all("/[[:upper:]#@w]/u", $str, $matches)) {
            $upperCount = count($matches[0]);
        }

        // width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256

        $widths = self::$fontWidths[$fontName] ?? self::$fontWidths['Calibri'];
        $n = ($len - $wideCount - $upperCount - $dotsCount) * $widths['lower'] +
            $wideCount * $widths['wide'] +
            $upperCount * $widths['upper'] +
            $dotsCount * $widths['dots'] + self::WIDTH_PADDING;

        $k = $fontSize / 11;

        return round($n * $k, 8);

    }

}