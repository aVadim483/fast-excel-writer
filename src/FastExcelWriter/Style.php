<?php

namespace avadim\FastExcelWriter;

/**
 * Class Style
 *
 * @package avadim\FastExcelWriter
 */
class Style
{
    public const FONT               = 'font';
    public const FONT_NAME          = 'name';
    public const FONT_STYLE         = 'style';
    public const FONT_STYLE_BOLD    = 'bold';
    public const FONT_STYLE_ITALIC  = 'italic';

    public const FONT_SIZE          = 'size';

    public const STYLE              = 'style';
    public const WIDTH              = 'width';

    public const TEXT_WRAP          = 'text-wrap';
    public const TEXT_ALIGN         = 'text-align';
    public const VERTICAL_ALIGN     = 'vertical-align';

    public const TEXT_ALIGN_LEFT    = 'left';
    public const TEXT_ALIGN_CENTER  = 'center';
    public const TEXT_ALIGN_RIGHT   = 'right';

    public const BORDER             = 'border';

    public const BORDER_SIDE        = 1;
    public const BORDER_STYLE       = 'style';
    public const BORDER_COLOR       = 'color';

    public const BORDER_TOP         = 1;
    public const BORDER_RIGHT       = 2;
    public const BORDER_BOTTOM      = 4;
    public const BORDER_LEFT        = 8;
    public const BORDER_ALL         = self::BORDER_TOP + self::BORDER_RIGHT + self::BORDER_BOTTOM + self::BORDER_LEFT;

    public const BORDER_NONE            = null;
    public const BORDER_THIN            = 'thin';
    public const BORDER_MEDIUM          = 'medium';
    public const BORDER_THICK           = 'thick';
    public const BORDER_DASH_DOT        = 'dashDot';
    public const BORDER_DASH_DOT_DOT    = 'dashDotDot';
    public const BORDER_DASHED          = 'dashed';
    public const BORDER_DOTTED          = 'dotted';
    public const BORDER_DOUBLE          = 'double';
    public const BORDER_HAIR            = 'hair';
    public const BORDER_MEDIUM_DASH_DOT = 'mediumDashDot';
    public const BORDER_MEDIUM_DASH_DOT_DOT = 'mediumDashDotDot';
    public const BORDER_MEDIUM_DASHED   = 'mediumDashed';
    public const BORDER_SLANT_DASH_DOT  = 'slantDashDot';

    public const BORDER_STYLE_MIN = self::BORDER_NONE;
    public const BORDER_STYLE_MAX = self::BORDER_SLANT_DASH_DOT;

    protected static $instance;

    protected static $fontStyleDefines = ['bold', 'italic', 'strike', 'underline'];

    /** @var array  */
    public $localeSettings = [];

    /** @var array  */
    public $defaultFont;

    /** @var array  */
    protected $cellStyles = [];

    /** @var array  */
    protected $numberFormats = [];

    protected $elements = [];

    protected $elementIndexes = [];


    /**
     * Constructor of Style
     *
     * @param $options
     */
    public function __construct($options)
    {
        self::$instance = $this;
        if (isset($options['default_font'])) {
            $this->setDefaultFont($options['default_font']);
        } else {
            $this->setDefaultFont(['name' => 'Arial', 'size' => 10]);
        }

        $defaultStyle = [
            'font' => $this->defaultFont,
            'fill' => 'none',
            'border' => 'none',
        ];
        $this->addCellStyle('GENERAL', $defaultStyle);
        $defaultStyle['fill'] = ['pattern' => 'gray125'];
        $this->addCellStyle('GENERAL', $defaultStyle);
    }

    /**
     * @param string $styleName
     *
     * @return string|null
     */
    protected static function _borderStyleName(string $styleName): ?string
    {
        static $styleNames = [
            self::BORDER_NONE => 0,
            self::BORDER_THIN => 1,
            self::BORDER_MEDIUM => 2,
            self::BORDER_THICK => 3,
            self::BORDER_DASH_DOT => 4,
            self::BORDER_DASH_DOT_DOT => 5,
            self::BORDER_DASHED => 6,
            self::BORDER_DOTTED => 7,
            self::BORDER_DOUBLE => 8,
            self::BORDER_HAIR => 9,
            self::BORDER_MEDIUM_DASH_DOT => 10,
            self::BORDER_MEDIUM_DASH_DOT_DOT => 11,
            self::BORDER_MEDIUM_DASHED => 12,
            self::BORDER_SLANT_DASH_DOT => 13,
        ];

        if (isset($styleNames[$styleName])) {
            return $styleName;
        }
        return null;
    }

    /**
     * @param array $font
     *
     * @return $this
     */
    public function setDefaultFont(array $font)
    {
        [$fontName, $fontFamily] = self::_getFamilyFont($font['name']);
        if ($fontFamily) {
            $font['name'] = $fontName;
            $font['family'] = $fontFamily;
        }
        $this->defaultFont = $font;

        return $this;
    }

    /**
     * @param $localeData
     *
     * @return $this
     */
    public function setLocaleSettings($localeData)
    {
        if (!empty($localeData['functions'])) {
            uksort($localeData['functions'], static function($a, $b) {
                return mb_strlen($b) - mb_strlen($a);
            });
        }
        if (!empty($localeData['formats'])) {
            uksort($localeData['formats'], static function($a, $b) {
                return mb_strlen($b) - mb_strlen($a);
            });
        }
        $this->localeSettings = $localeData;

        return $this;
    }

    /**
     * Examples:
     *  'thin' -> all sides are thin
     *  ['top' => ['style' => 'thin']]
     *  ['top' => ['style' => 'thin', 'color' => '#f00']]
     *
     * @param array|string $border
     *
     * @return array
     */
    public static function normalizeBorder($border): ?array
    {
        if (empty($border)) {
            return null;
        }

        if (is_scalar($border)) {
            if ($border[0] === '#') {
                // it's a color
                $border = ['all' => ['color' => $border]];
            } else {
                // it's a style
                $border = ['all' => ['style' => $border]];
            }
        }

        $result = [];
        if (is_array($border)) {
            /**
             * @var string|array $side
             * @var string|array $sideOptions
             */
            foreach($border as $side => $sideOptions) {
                $resultOptions = [];
                if ($sideOptions) {
                    if (is_array($sideOptions)) {
                        if (isset($sideOptions['style'])) {
                            $resultOptions['style'] = self::_borderStyleName($sideOptions['style']);
                        }
                        if (isset($sideOptions['color'])) {
                            $resultOptions['color'] = self::normalizeColor($sideOptions['color']);
                        }
                    } elseif ($sideOptions[0] === '#') {
                        $resultOptions['color'] = self::normalizeColor($sideOptions);
                    } else {
                        $resultOptions['style'] = self::_borderStyleName($sideOptions);
                    }
                }

                if (!is_numeric($side)) {
                    $side = strtolower($side);
                    if ($side === 'all') {
                        $side = self::BORDER_ALL;
                    }
                }
                if (is_numeric($side)) {
                    $side = (int)$side;
                    if ($side & self::BORDER_TOP) {
                        $result['top'] = isset($result['top']) ? array_merge($result['top'], $resultOptions) : $resultOptions;
                    }
                    if ($side & self::BORDER_RIGHT) {
                        $result['right'] = isset($result['right']) ? array_merge($result['right'], $resultOptions) : $resultOptions;
                    }
                    if ($side & self::BORDER_BOTTOM) {
                        $result['bottom'] = isset($result['bottom']) ? array_merge($result['bottom'], $resultOptions) : $resultOptions;
                    }
                    if ($side & self::BORDER_LEFT) {
                        $result['left'] = isset($result['left']) ? array_merge($result['left'], $resultOptions) : $resultOptions;
                    }
                } elseif ($side === 'top') {
                    $result['top'] = isset($result['top']) ? array_merge($result['top'], $resultOptions) : $resultOptions;
                } elseif ($side === 'right') {
                    $result['right'] = isset($result['right']) ? array_merge($result['right'], $resultOptions) : $resultOptions;
                } elseif ($side === 'bottom') {
                    $result['bottom'] = isset($result['bottom']) ? array_merge($result['bottom'], $resultOptions) : $resultOptions;
                } elseif ($side === 'left') {
                    $result['left'] = isset($result['left']) ? array_merge($result['left'], $resultOptions) : $resultOptions;
                }
            }
            self::_ksort($result);
        }
        return $result ?: null;
    }

    /**
     * @param array|string $fill
     *
     * @return array
     */
    public static function normalizeFill($fill): array
    {
        $result = [
            'patternFill' => ['_attributes' => ['patternType' => 'none']],
        ];
        if (!empty($fill) && $fill !== 'none') {
            $fillColor = null;
            if (!empty($fill['fill']) && is_string($fill['fill']) && $fill['fill'] !== 'none') {
                $fillColor = self::normalizeColor($fill['fill']);
            }
            elseif (!empty($fill['bg-color']) && $fill['bg-color'] !== 'none') {
                $fillColor = self::normalizeColor($fill['bg-color']);
            }
            elseif (!empty($fill['background-color']) && $fill['background-color'] !== 'none') {
                $fillColor = self::normalizeColor($fill['background-color']);
            }

            if ($fillColor) {
                $result['patternFill'] = [
                    '_attributes' => ['patternType' => 'solid'],
                    '_children' => [
                        'fgColor' => ['_attributes' => ['rgb' => $fillColor]],
                        'bgColor' => ['_attributes' => ['indexed' => 64]],
                    ],
                ];
            }
            if (!empty($fill['pattern'])) {
                $result['patternFill']['_attributes']['patternType'] = $fill['pattern'];
            }
            self::_ksort($result);
        }

        return $result;
    }

    /**
     * @param string $color
     *
     * @return string|null
     */
    public static function normalizeColor(string $color): ?string
    {
        if ($color) {
            if (strpos($color, '#') === 0) {
                $color = substr($color, 1, 6);
            }
            $color = strtoupper($color);
            if (preg_match('/^[0-9A-F]+$/i', $color)) {
                if (strlen($color) === 3) {
                    $color = $color[0] . $color[0] . $color[1] . $color[1] . $color[2] . $color[2];
                }
                if (strlen($color) === 6) {
                    $color = 'FF' . $color;
                }
                if (strlen($color) > 8) {
                    return substr($color, 1, 8);
                }
                if (strlen($color) === 8) {
                    return $color;
                }
            }
        }
        return null;
    }

    /**
     * @param $fontName
     *
     * @return array|null[]
     */
    protected static function _getFamilyFont($fontName): array
    {
        $defaultFonts = [
            'Times New Roman' => [
                'name' => 'Times New Roman',
                'family' => 1,
            ],
            'Arial' => [
                'name' => 'Arial',
                'family' => 2,
            ],
            'Courier New' => [
                'name' => 'Courier New',
                'family' => 3,
            ],
            'Comic Sans MS' => [
                'name' => 'Comic Sans MS',
                'family' => 4,
            ],
        ];

        foreach ($defaultFonts as $name => $defFont) {
            if (strcasecmp($fontName, $name) === 0) {
                return [$defFont['name'], $defFont['family']];
            }
        }
        return [null, null];
    }

    /**
     * @param array|string $font
     *
     * @return array
     */
    public static function normalizeFont($font): array
    {
        $result = self::$instance->defaultFont;

        if (!empty($font)) {
            if (is_string($font)) {
                if (in_array($font, self::$fontStyleDefines, true)) {
                    $font = ['style' => $font];
                }
                else {
                    $font = [];
                }
            }
            foreach($font as $key => $val) {
                switch ($key) {
                    case 'name':
                        [$fontName, $fontFamily] = self::_getFamilyFont($font['name']);
                        if ($fontFamily) {
                            $result['name'] = $fontName;
                            $result['family'] = $fontFamily;
                        }
                        break;
                    case 'style':
                        if (is_string($val)) {
                            if (strpos($val, 'bold') !== false) {
                                $result['style-bold'] = 1;
                            }
                            if (strpos($val, 'italic') !== false) {
                                $result['style-italic'] = 1;
                            }
                            if (strpos($val, 'strike') !== false) {
                                $result['style-strike'] = 1;
                            }
                            if (strpos($val, 'underline') !== false) {
                                $result['style-underline'] = 1;
                            }
                        }
                        break;
                    case 'size':
                        $result['size'] = (float)$val;
                        break;
                    case 'color':
                        $result['color'] = ['rgb' => self::normalizeColor($val)];
                        break;
                    default:
                        $result[$key] = $val;
                }
            }
            self::_ksort($result);
        }
        return $result;
    }

    /**
     * @param $style
     *
     * @return array
     */
    public static function normalize($style): array
    {
        $result = [];
        if (is_array($style)) {
            foreach($style as $styleKey => $styleVal) {
                switch ($styleKey) {
                    case 'border':
                        $result[$styleKey] = self::normalizeBorder($styleVal);
                        break;
                    case 'color':
                    case 'text-color':
                    case 'font-color':
                        $result['color'] = $styleVal;
                        break;
                    case 'fill':
                    case 'bg-color':
                    case 'background-color':
                    case 'cell-color':
                        $result['fill'] = $styleVal;
                        break;
                    case 'font':
                        $result['font'] = self::normalizeFont($styleVal);
                        break;
                    case 'text-align':
                    case 'align':
                    case 'halign':
                        if (in_array($styleVal, ['general', 'left', 'right', 'justify', 'center'])) {
                            $result['text-align'] = $styleVal;
                        }
                        break;
                    case 'vertical-align':
                    case 'valign':
                        if (in_array($styleVal, ['bottom', 'center', 'distributed', 'top'])) {
                            $result['vertical-align'] = $styleVal;
                        }
                        break;
                    case 'text-wrap':
                        $result['text-wrap'] = (bool)$styleVal;
                        break;
                    default:
                        $result[$styleKey] = $styleVal;
                }
            }
        }
        return $result;
    }

    /**
     * @param array $array
     */
    protected static function _ksort(array &$array)
    {
        if ($array) {
            ksort($array);
            foreach($array as $key => $val) {
                if (is_array($val)) {
                    self::_ksort($val);
                    $array[$key] = $val;
                }
            }
        }
    }

    /**
     * @param $haystack
     * @param $needle
     *
     * @return int
     */
    public static function addToListGetIndex(&$haystack, $needle)
    {
        $existingIdx = array_search($needle, $haystack, $strict = true);
        if ($existingIdx === false) {
            $existingIdx = count($haystack);
            $haystack[] = $needle;
        }
        return $existingIdx;
    }

    /**
     * @param string $sectionName
     * @param int $index
     *
     * @return array
     */
    protected function findElement(string $sectionName, int $index)
    {
        if (isset($this->elementIndexes[$index], $this->elements[$sectionName][$this->elementIndexes[$index]])) {
            return $this->elements[$sectionName][$this->elementIndexes[$index]];
        }

        return [];
    }

    /**
     * @param string $sectionName
     * @param string|array $value
     *
     * @return int
     */
    protected function addElement(string $sectionName, $value)
    {
        $key = json_encode($value);
        if (isset($this->elements[$sectionName][$key])) {
            return $this->elements[$sectionName][$key]['index'];
        }
        $index = empty($this->elements[$sectionName]) ? 0 : count($this->elements[$sectionName]);
        $this->elements[$sectionName][$key] = [
            'index' => $index,
            'value' => $value,
        ];
        $this->elementIndexes[$index] = $key;

        return $index;
    }

    /**
     * @param array $cellStyle
     * @param array $fullStyle
     */
    protected function addStyleFont(array &$cellStyle, array &$fullStyle = [])
    {
        $index = 0;
        if (isset($cellStyle['font']) || isset($cellStyle['color']) || isset($cellStyle['text-color']) || isset($cellStyle['font-style']) || isset($cellStyle['font-size'])) {
            if (!isset($cellStyle['font'])) {
                $cellStyle['font'] = [];
            }
            elseif (is_string($cellStyle['font'])) {
                if (in_array($cellStyle['font'], self::$fontStyleDefines, true)) {
                    $cellStyle['font'] = ['style' => $cellStyle['font']];
                }
                else {
                    $cellStyle['font'] = [];
                }
            }

            if (!empty($cellStyle['color'])) {
                $cellStyle['font']['color'] = $cellStyle['color'];
                unset($cellStyle['color']);
            }
            elseif (!empty($cellStyle['text-color'])) {
                $cellStyle['font']['color'] = $cellStyle['text-color'];
                unset($cellStyle['text-color']);
            }
            elseif (!empty($cellStyle['font-color'])) {
                $cellStyle['font']['color'] = $cellStyle['font-color'];
                unset($cellStyle['font-color']);
            }
            if (!empty($cellStyle['font-style']) && empty($cellStyle['font']['style'])) {
                $cellStyle['font']['style'] = $cellStyle['font-style'];
                unset($cellStyle['font-style']);
            }
            if (!empty($cellStyle['font-size']) && empty($cellStyle['font']['size'])) {
                $cellStyle['font']['size'] = $cellStyle['font-size'];
                unset($cellStyle['font-size']);
            }

            $value = self::normalizeFont($cellStyle['font']);
            $index = $this->addElement('fonts', $value);

            if (isset($cellStyle['font'])) {
                unset($cellStyle['font']);
            }

            $fullStyle['font'] = $value;
        }
        else {
            $fullStyle['font'] = $this->findElement('fonts', $index);
        }
        $cellStyle['fontId'] = $index;
    }

    /**
     * @param array $cellStyle
     * @param array $fullStyle
     */
    protected function addStyleFill(array &$cellStyle, array &$fullStyle = [])
    {
        $index = 0;
        $fill = [];
        if (isset($cellStyle['fill'])) {
            if (is_array($cellStyle['fill'])) {
                $fill = $cellStyle['fill'];
            }
            else {
                $fill['fill'] = $cellStyle['fill'];
            }
            unset($cellStyle['fill']);
        }
        elseif (!empty($cellStyle['bg-color'])) {
            $fill['fill'] = $cellStyle['bg-color'];
            unset($cellStyle['bg-color']);
        }
        elseif (!empty($cellStyle['background-color'])) {
            $fill['fill'] = $cellStyle['background-color'];
            unset($cellStyle['background-color']);
        }

        if (isset($cellStyle['color'])) {
            $fill['color'] = $cellStyle['color'];
            unset($cellStyle['color']);
        }
        elseif (!empty($cellStyle['fg-color'])) {
            $fill['color'] = $cellStyle['fg-color'];
            unset($cellStyle['fg-color']);
        }

        if ($fill) {
            $value = self::normalizeFill($fill);
            $index = $this->addElement('fills', $value);

            $fullStyle['fills'] = $value;
        }
        else {
            $fullStyle['fill'] = $this->findElement('fills', $index);
        }
        $cellStyle['fillId'] = $index;
    }

    /**
     * @param array $cellStyle
     * @param array $fullStyle
     */
    protected function addStyleBorder(array &$cellStyle, array &$fullStyle = [])
    {
        $index = 0;
        if (isset($cellStyle['border'])) {
            if ($cellStyle['border']) {
                $value = self::normalizeBorder($cellStyle['border']);
                $index = $this->addElement('borders', $value);

                $fullStyle['borders'] = $value;
            }
            else {
                $fullStyle['border'] = $this->findElement('borders', $index);
            }
            unset($cellStyle['border']);
        }
        $cellStyle['borderId'] = $index;
    }

    /**
     * @param array $cellStyle
     *
     * @return int
     */
    protected function indexStyle($cellStyle)
    {
        self::_ksort($cellStyle);

        return $this->addElement('cellXfs', $cellStyle);
    }

    /**
     * @param string $numberFormat
     * @param array|null $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    public function addCellStyle(string $numberFormat, ?array $cellStyle = [], ?array &$fullStyle = [])
    {
        $fullStyle = [];
        if (empty($cellStyle)) {
            $cellStyle = [];
        }
        $this->addStyleFont($cellStyle, $fullStyle);
        $this->addStyleFill($cellStyle, $fullStyle);
        $this->addStyleBorder($cellStyle, $fullStyle);
        $cellStyle['numFmtId'] = self::addToListGetIndex($this->numberFormats, $numberFormat);

        return $this->indexStyle($cellStyle);
    }

    /**
     * @param string $sectionName
     *
     * @return array
     */
    protected function getElements(string $sectionName)
    {
        if (!empty($this->elements[$sectionName])) {
            $result = [];
            foreach ($this->elements[$sectionName] as $element) {
                $result[$element['index']] = $element['value'];
            }
            return $result;
        }

        return [];
    }

    /**
     * @return array
     */
    public function getStyleFonts()
    {
        return $this->getElements('fonts');
    }

    /**
     * @return array
     */
    public function getStyleFills()
    {
        return $this->getElements('fills');
    }

    /**
     * @return array
     */
    public function getStyleBorders()
    {
        return $this->getElements('borders');
    }

    /**
     * @return array
     */
    public function getStyleCellXfs()
    {
        return $this->getElements('cellXfs');
    }

    /**
     * @param $numFormat
     *
     * @return string
     */
    private static function determineNumberFormatType($numFormat)
    {
        if ($numFormat === 'GENERAL') {
            return 'n_auto';
        }
        if ($numFormat === '@') {
            return 'n_string';
        }
        if ($numFormat === '0') {
            return 'n_numeric';
        }
        if (preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_datetime';
        }
        if (preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_datetime';
        }
        if (preg_match('/[Y]{2,4}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        if (preg_match('/[D]{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        if (preg_match('/[M]{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        if (preg_match('/\$(?![^"]*+")/', $numFormat)) {
            return 'n_numeric';
        }
        if (preg_match('/%(?![^"]*+")/', $numFormat)) {
            return 'n_numeric';
        }
        if (preg_match('/0(?![^"]*+")/', $numFormat)) {
            return 'n_numeric';
        }
        return 'n_auto';
    }

    /**
     * @param $numFormat
     *
     * @return string
     */
    private static function numberFormatStandardized($numFormat)
    {
        $stack = [];
        if (!is_scalar($numFormat) || $numFormat === 'auto' || $numFormat === '' || $numFormat === 'GENERAL') {
            return 'GENERAL';
        }
        if ($numFormat === 'string' || $numFormat === 'text') {
            return '@';
        }
        if ($numFormat === 'integer' || $numFormat === 'int') {
            return '0';
        }
        if ($numFormat === 'percent') {
            return '0%';
        }
        while (isset(self::$instance->localeSettings['formats'][$numFormat])) {
            if (!$numFormat || isset($stack[$numFormat])) {
                break;
            }
            if (isset(self::$instance->localeSettings['formats'][$numFormat])) {
                $numFormat = self::$instance->localeSettings['formats'][$numFormat];
            } else {
                break;
            }
        }

        $ignoreUntil = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($numFormat); $i < $ix; $i++) {
            $c = $numFormat[$i];

            if ($ignoreUntil === '' && $c === '[') {
                $ignoreUntil = ']';
            } elseif ($ignoreUntil === '' && $c === '"') {
                $ignoreUntil = '"';
            } elseif ($ignoreUntil === $c) {
                $ignoreUntil = '';
            }

            if ($ignoreUntil === '' && ($c === ' ' || $c === '-' || $c === '(' || $c === ')') && ($i === 0 || $numFormat[$i - 1] !== '_')) {
                $escaped .= "\\" . $c;
            } else {
                $escaped .= $c;
            }
        }
        return $escaped;
    }

    /**
     * @param $format
     *
     * @return array
     */
    public function defineFormatType($format)
    {
        if (is_array($format)) {
            $format = reset($format);
        }
        $numberFormat = self::numberFormatStandardized($format);
        $numberFormatType = self::determineNumberFormatType($numberFormat);
        $cellStyleIdx = $this->addCellStyle($numberFormat, null);

        $formatType = [
            'number_format' => $numberFormat, //contains excel format like 'YYYY-MM-DD HH:MM:SS'
            'number_format_type' => $numberFormatType, //contains friendly format like 'datetime'
            'default_style_idx' => $cellStyleIdx,
        ];

        return $formatType;
    }

    /**
     * @return array
     */
    public function defaultFormatType()
    {
        static $defaultFormatType;

        if (!$defaultFormatType) {
            $defaultFormatType = $this->defineFormatType('GENERAL');
        }
        return $defaultFormatType;
    }

    /**
     * @return array
     */
    /*
    public function _styleFontIndexes()
    {
        $fills = ['', ''];  // 2 placeholders for static xml later
        $fonts = ['', ''];  // 2 placeholders for static xml later
        $borders = [''];    // 1 placeholder for static xml later
        $styleIndexes = [];
        foreach ($this->cellStyles as $i => $cellStyleString) {
            $semiColonPos = strpos($cellStyleString, ";");
            $numberFormatIdx = substr($cellStyleString, 0, $semiColonPos);
            $styleJsonString = substr($cellStyleString, $semiColonPos + 1);
            $style = json_decode($styleJsonString, true);

            $styleIndexes[$i] = ['num_fmt_idx' => $numberFormatIdx];//initialize entry

            // new border settings
            if (!empty($style['border']) && is_array($style['border'])) {
                $borderValue = [];
                foreach($style['border'] as $side => $options) {
                    $borderValue[$side] = $options;
                    if (!empty($options['color'])) {
                        $color = Style::normaliazeColor($options['color']);
                        if ($color) {
                            $borderValue[$side]['color'] = $color;
                        }
                    }
                }
                $styleIndexes[$i]['border_idx'] = self::addToListGetIndex($borders, $borderValue);
            }
            if (!empty($style['fill'])) {
                $color = Style::normaliazeColor($style['fill']);
                if ($color) {
                    $styleIndexes[$i]['fill_idx'] = self::addToListGetIndex($fills, $color);
                }
            }
            if (!empty($style['text-align'])) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['text-align'] = $style['text-align'];
            }
            if (!empty($style['vertical-align'])) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['vertical-align'] = $style['vertical-align'];
            }
            if (!empty($style['text-wrap'])) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['text-wrap'] = true;
            }

            $font = null;
            if (!empty($style['font'])) {
                $font = Style::normalizeFont($style['font']);
            }
            if (!$font) {
                $font = Style::normalizeFont([]);
            }
            if (isset($style['color'])) {
                $color = Style::normaliazeColor($style['color']);
                if ($color) {
                    $font['color'] = $color;
                }
            }
            $styleIndexes[$i]['font_idx'] = self::addToListGetIndex($fonts, $font);
        }
        return ['fills' => $fills, 'fonts' => $fonts, 'borders' => $borders, 'styles' => $styleIndexes];
    }
    */

    /**
     * @return array
     */
    public function _getNumberFormats()
    {
        return $this->numberFormats;
    }
}

// EOF