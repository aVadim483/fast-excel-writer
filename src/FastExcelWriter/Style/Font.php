<?php

namespace avadim\FastExcelWriter\Style;

use avadim\FastExcelWriter\RichText\RichText;

class Font
{
    public const DEFAULT_FONT_NAME = 'Calibri';
    public const DEFAULT_FONT_SIZE = 11;

    protected static float $widthFactor = 1.0;
    protected WidthCache $widthCache;
    protected WidthCache $metricCache;
    protected DisplayFormatter $formatter;
    protected array $lastBaseFont = [];
    protected ?float $lastDigitWidth = null;
    protected array $simpleFont = [];
    protected array $simplePixels = [];

    public function __construct()
    {
        $this->widthCache = new WidthCache();
        $this->metricCache = new WidthCache(32, 8192, 256);
        $this->formatter = new DisplayFormatter();
    }

    public static function setFontWidthFactor(float $factor)
    {
        self::$widthFactor = $factor;
    }

    public static function calcTextWidth(string $fontName, $fontSize, string $value, ?string $numberFormat = null): float
    {
        static $calculator;
        if (!$calculator) {
            $calculator = new self();
        }
        return $calculator->calculateTextWidth($fontName, $fontSize, $value, $numberFormat);
    }

    public function calculateTextWidth(string $fontName, $fontSize, string $value, ?string $numberFormat = null): float
    {
        $format = $numberFormat ?? '';
        $key = 'F' . strlen($fontName) . ':' . $fontName . ':' . (float)$fontSize . ':' . strlen($format) . ':' . $format . $value;
        $pixels = $this->widthCache->get($key);
        if ($pixels === null) {
            if ($numberFormat !== null && $numberFormat !== '@' && strcasecmp($numberFormat, 'general') !== 0) {
                $value = $this->formatter->format($value, $numberFormat, is_numeric($value));
            }
            $pixels = $this->measurePixels($value, ['name' => $fontName, 'size' => $fontSize]);
            $this->widthCache->put($key, $pixels);
        }
        return $this->excelWidth($pixels, []);
    }

    public function formatDisplayValue($value, string $format, bool $numeric, string $locale): string
    {
        return $this->formatter->format($value, $format, $numeric, $locale);
    }

    public function dateWidthSample(string $format): ?string
    {
        return $this->formatter->dateWidthSample($format);
    }

    /** Cache signed decimal lengths for one active font; retain unscaled pixels. */
    public function calculateIntegerWidth(int $value, array $font, array $baseFont): float
    {
        if ($font !== $this->simpleFont) {
            $this->simpleFont = $font;
            $this->simplePixels = [];
        }
        $text = (string)$value;
        $length = strlen($text);
        $shape = $value < 0 ? -$length : $length;
        if (!isset($this->simplePixels[$shape])) {
            $this->simplePixels[$shape] = $this->measurePixels($text, $font);
        }
        return $this->excelWidth($this->simplePixels[$shape], $baseFont);
    }

    /** Boolean literals ignore number masks and need only two measurements per font. */
    public function calculateBooleanWidth(bool $value, array $font, array $baseFont): float
    {
        if ($font !== $this->simpleFont) {
            $this->simpleFont = $font;
            $this->simplePixels = [];
        }
        $text = $value ? 'TRUE' : 'FALSE';
        if (!isset($this->simplePixels[$text])) {
            $this->simplePixels[$text] = $this->measurePixels($text, $font);
        }
        return $this->excelWidth($this->simplePixels[$text], $baseFont);
    }

    /** Content measurements have no padding, workbook scaling or global factor. */
    protected function measurePixels(string $text, array $font): float
    {
        $name = $font['name'] ?? self::DEFAULT_FONT_NAME;
        $size = (float)($font['size'] ?? self::DEFAULT_FONT_SIZE);
        $face = (!empty($font['style-bold']) ? 1 : 0) + (!empty($font['style-italic']) ? 2 : 0);
        // All supported metric tables use tabular digits. Normalize them only in the
        // measurement cache key, after value formatting; the saved text is untouched.
        $key = 'T' . strlen($name) . ':' . $name . ':' . $size . ':' . $face . ':' . strtr($text, '123456789', '000000000');
        $width = $this->widthCache->get($key);
        if ($width !== null) {
            return $width;
        }
        $metricKey = strtolower($name) . ':' . $face;
        $metrics = $this->metricCache->get($metricKey);
        if ($metrics === null) {
            $names = ['calibri' => 'Calibri', 'arial' => 'Arial', 'tahoma' => 'Tahoma', 'times new roman' => 'Times New Roman'];
            $table = FontMetrics::WIDTHS[$names[strtolower($name)] ?? self::DEFAULT_FONT_NAME][$face];
            $widths = array_combine(mb_str_split(FontMetrics::CHARACTERS, 1, 'UTF-8'), $table);
            $translation = [];
            foreach ($widths as $character => $advance) {
                $translation[(string)$character] = chr($advance);
            }
            for ($code = 0; $code < 32; ++$code) {
                $translation[chr($code)] = chr(7);
            }
            $translation[chr(127)] = chr(7);
            $metrics = ['widths' => $widths, 'translation' => $translation];
            $this->metricCache->put($metricKey, $metrics);
        }
        if (strlen($text) >= 128) {
            $advances = strtr($text, $metrics['translation']);
            if (!preg_match('/[\x80-\xff]/', $advances)) {
                $width = 0;
                foreach (count_chars($advances, 1) as $advance => $count) {
                    $width += $advance * $count;
                }
                $width *= $size / self::DEFAULT_FONT_SIZE;
                $this->widthCache->put($key, $width);
                return $width;
            }
        }
        $width = 0;
        $widths = $metrics['widths'];
        if (preg_match('/[\x80-\xff]/', $text)) {
            $characters = mb_str_split(mb_convert_encoding($text, 'UTF-8', 'UTF-8'), 1, 'UTF-8');
        }
        else {
            $characters = str_split($text);
        }
        foreach ($characters as $character) {
            if (isset($widths[$character])) {
                $width += $widths[$character];
            }
            else {
                $code = mb_ord($character);
                if (($code >= 0x300 && $code <= 0x36f) || $code === 0x200d || ($code >= 0xfe00 && $code <= 0xfe0f)) {
                    continue;
                }
                $width += ($code >= 0x2e80 && $code <= 0xa4cf) || $code >= 0x1f000 ? 14 : 7;
            }
        }
        $width *= $size / self::DEFAULT_FONT_SIZE;
        $this->widthCache->put($key, $width);
        return $width;
    }

    protected function excelWidth(float $pixels, array $baseFont): float
    {
        if ($this->lastDigitWidth !== null && $baseFont === $this->lastBaseFont) {
            return $this->scaleWidth(($pixels + 5) / max(1, $this->lastDigitWidth));
        }
        $key = 'digits:' . json_encode([$baseFont['name'] ?? self::DEFAULT_FONT_NAME,
            $baseFont['size'] ?? self::DEFAULT_FONT_SIZE, !empty($baseFont['style-bold']), !empty($baseFont['style-italic'])]);
        $maximumDigit = $this->metricCache->get($key);
        if ($maximumDigit === null) {
            $maximumDigit = 0.0;
            foreach (range(0, 9) as $digit) {
                $maximumDigit = max($maximumDigit, $this->measurePixels((string)$digit, $baseFont));
            }
            $this->metricCache->put($key, $maximumDigit);
        }
        $this->lastBaseFont = $baseFont;
        $this->lastDigitWidth = $maximumDigit;
        // OOXML widths include five pixels of padding and are stored in 1/256 units.
        return $this->scaleWidth(($pixels + 5) / max(1, $maximumDigit));
    }

    protected function scaleWidth(float $width): float
    {
        return min(255.0, max(0.0, floor($width * self::$widthFactor * 256) / 256));
    }

    public function calculateRichTextWidth(RichText $value, array $font, bool $wrap): float
    {
        return $this->calculateCellWidth($value, $font, [], $wrap);
    }

    /** Measure the longest visual line; add padding once for the whole cell. */
    public function calculateCellWidth($value, array $font, array $baseFont, bool $wrap): float
    {
        if (!$value instanceof RichText && (!$wrap || strpos((string)$value, "\n") === false)) {
            return $this->excelWidth($this->measurePixels((string)$value, $font), $baseFont);
        }
        $lineWidth = 0.0;
        $maxWidth = 0.0;
        $runs = [];
        if ($value instanceof RichText) {
            foreach ($value->fragments() as $fragment) {
                $properties = $fragment->getProperties();
                $runFont = $font;
                foreach (['f' => 'name', 'sz' => 'size', 'b' => 'style-bold', 'i' => 'style-italic'] as $source => $target) {
                    if ($properties[$source] !== null) {
                        $runFont[$target] = $properties[$source];
                    }
                }
                if (in_array($properties['vertAlign'], ['subscript', 'superscript'], true)) {
                    $runFont['size'] = ($runFont['size'] ?? self::DEFAULT_FONT_SIZE) * 0.65;
                }
                $runs[] = [$fragment->getText(), $runFont];
            }
        }
        else {
            $runs[] = [(string)$value, $font];
        }
        foreach ($runs as $run) {
            $text = str_replace("\r\n", "\n", $run[0]);
            $lines = $wrap ? explode("\n", $text) : [$text];
            foreach ($lines as $index => $line) {
                if ($index) {
                    $maxWidth = max($maxWidth, $lineWidth);
                    $lineWidth = 0.0;
                }
                $lineWidth += $this->measurePixels($line, $run[1]);
            }
        }
        return $this->excelWidth(max($maxWidth, $lineWidth), $baseFont);
    }

    /** Kept for subclasses using the previous protected helper. */
    protected static function _calcWidth(string $str, string $fontName, $fontSize, ?bool $numFormat = false): float
    {
        $calculator = new self();
        return $calculator->calculateTextWidth($fontName, $fontSize, $str);
    }
}
