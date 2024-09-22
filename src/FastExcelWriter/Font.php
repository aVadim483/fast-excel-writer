<?php

namespace avadim\FastExcelWriter;

class Font
{
    public const DEFAULT_FONT_NAME = 'Calibri';
    public const DEFAULT_FONT_SIZE = 11;

    // constants for auo width
    protected const WIDTH_PADDING = 0.81;
    protected const DEFAULT_FONT_RANGE = 'default';
    protected const DEFAULT_FONT_WIDTH = 1.2;


    protected static array $fontWidths = [
        'Aptos' => [
            'default' => 1.07,
            'dots' => 0.55,
            'upper' => 1.48,
            'wide' => 1.76,
            'arabic' => 1.00,
            'hebrew' => 0.80,
            'chinese' => 2.18
        ],
        'Arial' => [
            'default' => 1.10,
            'dots' => 0.50,
            'upper' => 1.35,
            'wide' => 1.77,
            'arabic' => 1.00,
            'hebrew' => 0.92,
            'chinese' => 2.18
        ],
        'Calibri' => [
            'default' => 1.07,
            'dots' => 0.60,
            'upper' => 1.21,
            'wide' => 1.72,
            'arabic' => 1.00,
            'hebrew' => 0.86,
            'chinese' => 2.18
        ],
        'Tahoma' => [
            'default' => 1.11,
            'dots' => 0.68,
            'upper' => 1.22,
            'wide' => 1.89,
            'arabic' => 0.76,
            'hebrew' => 1.10,
            'chinese' => 2.20
        ],
        'Times New Roman' => [
            'default' => 1.00,
            'dots' => 0.58,
            'upper' => 1.48,
            'wide' => 1.91,
            'arabic' => 0.60,
            'hebrew' => 1.12,
            'chinese' => 2.20
        ],
    ];

    protected static float $widthFactor = 1.0;


    public static function setFontWidthFactor(float $factor)
    {
        self::$widthFactor = $factor;
    }

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

    protected static function _ranges(): array
    {
        static $ranges = [];

        if (!$ranges) {
            $ranges = [
                'dots' => "/[,.\-:'\";`IiJjfl\[\]\(\)\{\}]/u",
                'wide' => "/[@%&WMQ]/u",
                'upper' => "/[[:upper:]#@w]/u",
                'hebrew' => '/[' . mb_chr(0x0590) . '-' . mb_chr(0x05FF) . ']/u',
                'arabic' => '/[' . mb_chr(0x0600) . '-' . mb_chr(0x06FF) . ']/u',
                'chinese' => '/[' . mb_chr(0x4e00) . '-' . mb_chr(0x9FCC) . ']/u',
            ];
        }
        return $ranges;
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
        $strWidth = 0;
        foreach (self::_ranges() as $range => $pattern) {
            if (preg_match_all($pattern, $str, $matches)) {
                $count = count($matches[0]);
                if (isset(self::$fontWidths[$fontName][$range])) {
                    $charWidth = self::$fontWidths[$fontName][$range];
                }
                elseif (isset(self::$fontWidths[self::DEFAULT_FONT_NAME][$range])) {
                    $charWidth = self::$fontWidths[self::DEFAULT_FONT_NAME][$range];
                }
                elseif (isset(self::$fontWidths[self::DEFAULT_FONT_NAME][self::DEFAULT_FONT_RANGE])) {
                    $charWidth = self::$fontWidths[self::DEFAULT_FONT_NAME][self::DEFAULT_FONT_RANGE];
                }
                else {
                    $charWidth = self::DEFAULT_FONT_WIDTH;
                }
                $strWidth += $count * $charWidth;
                $len -= $count;
                $str = preg_replace($pattern, '', $str);
            }
        }
        if (isset(self::$fontWidths[$fontName][self::DEFAULT_FONT_RANGE])) {
            $charWidth = self::$fontWidths[$fontName][self::DEFAULT_FONT_RANGE];
        }
        elseif (isset(self::$fontWidths[self::DEFAULT_FONT_NAME][self::DEFAULT_FONT_RANGE])) {
            $charWidth = self::$fontWidths[self::DEFAULT_FONT_NAME][self::DEFAULT_FONT_RANGE];
        }
        else {
            $charWidth = self::DEFAULT_FONT_WIDTH;
        }
        $n = $strWidth + $len * $charWidth + self::WIDTH_PADDING;
        $k = $fontSize / self::DEFAULT_FONT_SIZE * self::$widthFactor;

        return round($n * $k, 8);

    }

}