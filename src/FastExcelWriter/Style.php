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

    public const BORDER_LEFT        = 1;
    public const BORDER_RIGHT       = 2;
    public const BORDER_TOP         = 4;
    public const BORDER_BOTTOM      = 8;
    public const BORDER_ALL         = self::BORDER_TOP + self::BORDER_RIGHT + self::BORDER_BOTTOM + self::BORDER_LEFT;

    public const BORDER_NONE = null;
    public const BORDER_THIN = 'thin';
    public const BORDER_MEDIUM = 'medium';
    public const BORDER_THICK = 'thick';
    public const BORDER_DASH_DOT = 'dashDot';
    public const BORDER_DASH_DOT_DOT = 'dashDotDot';
    public const BORDER_DASHED = 'dashed';
    public const BORDER_DOTTED = 'dotted';
    public const BORDER_DOUBLE = 'double';
    public const BORDER_HAIR = 'hair';
    public const BORDER_MEDIUM_DASH_DOT = 'mediumDashDot';
    public const BORDER_MEDIUM_DASH_DOT_DOT = 'mediumDashDotDot';
    public const BORDER_MEDIUM_DASHED = 'mediumDashed';
    public const BORDER_SLANT_DASH_DOT = 'slantDashDot';

    public const BORDER_STYLE_MIN = self::BORDER_NONE;
    public const BORDER_STYLE_MAX = self::BORDER_SLANT_DASH_DOT;

    protected static $instance;

    protected static array $fontStyleDefines = ['bold', 'italic', 'strike', 'underline'];

    public array $localeSettings = [];

    public array $defaultFont = [];
    protected int $defaultFontId = -1;

    public array $defaultStyle = [];

    /** @var array Specified styles for hyperlinks */
    public array $hyperlinkStyle = [];

    /** @var array Specified styles for formats '@...'  */
    public array $defaultFormatStyles = [];

    protected array $elements = [];

    protected array $elementIndexes = [];


    /**
     * Constructor of Style
     *
     * @param array|null $options
     */
    public function __construct(?array $options)
    {
        self::$instance = $this;
        $defaultFont = [
            'name' => 'Arial',
            'size' => 10,
        ];
        $defaultStyle = [
            'font' => $this->defaultFont,
            'fill' => 'none',
            'border' => 'none',
        ];
        $defaultFormatStyles = [];
        $hyperlinkStyle = [
            'font' => ['color' => '0563C1', 'underline' => true],
        ];

        if (isset($options['default_font'])) {
            foreach($options['default_font'] as $key => $font) {
                $key = strtoupper($key);
                if (isset($defaultFont[$key])) {
                    $defaultFont[$key] = array_merge($defaultFont[$key], $font);
                }
            }
        }

        $this->setDefaultFont($defaultFont);
        $this->setDefaultStyle($defaultStyle);
        $this->addCellStyle('GENERAL', $defaultStyle);

        $defaultStyle['fill'] = ['pattern' => 'gray125'];
        $this->addCellStyle('GENERAL', $defaultStyle);

        $this->hyperlinkStyle = $hyperlinkStyle;
        $this->defaultFormatStyles = $defaultFormatStyles;
    }

    /**
     * @param array $styles
     *
     * @return array
     */
    public static function mergeStyles(array $styles)
    {
        $result = [];
        if ($styles) {
            $set = array_filter($styles);
            if ($set) {
                if (count($set) === 1) {
                    $result = reset($set);
                }
                else {
                    $result = array_replace_recursive(...$set);
                }

            }
        }
        return $result;
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
        $this->defaultFont = self::normalizeFont($font);

        return $this;
    }

    /**
     * @param array $style
     *
     * @return $this
     */
    public function setDefaultStyle(array $style)
    {
        $this->defaultStyle = $style;

        return $this;
    }

    /**
     * @param array $localeData
     *
     * @return $this
     */
    public function setLocaleSettings(array $localeData)
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

    public static function borderOptions($border)
    {
        $result = [];
        if (is_scalar($border)) {
            if ($border[0] === '#') {
                // it's a color
                $result = [Style::BORDER_ALL => ['color' => $border]];
            }
            else {
                // it's a style
                $result = [Style::BORDER_ALL => ['style' => $border]];
            }
        }
        else {
            foreach($border as $side => $sideOptions) {
                if (!is_numeric($side)) {
                    switch (strtolower($side)) {
                        case 'all':
                            $result[self::BORDER_ALL] = $sideOptions;
                            break;
                        case 'left':
                            $result[self::BORDER_LEFT] = $sideOptions;
                            break;
                        case 'right':
                            $result[self::BORDER_RIGHT] = $sideOptions;
                            break;
                        case 'top':
                            $result[self::BORDER_TOP] = $sideOptions;
                            break;
                        case 'bottom':
                            $result[self::BORDER_BOTTOM] = $sideOptions;
                            break;
                        default:
                            // nothing
                    }
                }
                elseif ($side & self::BORDER_ALL) {
                    $result[$side] = $sideOptions;
                }
            }
        }

        return $result;
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
        $result = [
            'val' => [
                'left' => ['style' => null, 'color' => null],
                'right' => ['style' => null, 'color' => null],
                'top' => ['style' => null, 'color' => null],
                'bottom' => ['style' => null, 'color' => null],
            ],
            'tag' => [
                'border' => '<border diagonalDown="false" diagonalUp="false">',
                'left' => '<left/>',
                'right' => '<right/>',
                'top' => '<top/>',
                'bottom' => '<bottom/>',
                'diagonal' => '<diagonal/>',
                '/border' => '</border>',
            ],
        ];

        if (!empty($border) && $border !== 'none') {
            $border = self::borderOptions($border);
            foreach($border as $side => $sideOptions) {
                if ($side & self::BORDER_LEFT) {
                    $style = isset($sideOptions['style']) ? self::_borderStyleName($sideOptions['style']) : null;
                    $color = isset($sideOptions['color']) ? self::normalizeColor($sideOptions['color']) : null;
                    if ($style || $color) {
                        if ($style) {
                            $result['val']['left']['style'] = $style;
                        }
                        $result['tag']['left'] = '<left' . ($style ? ' style="' . $style . '"' : '');
                        if ($color) {
                            $result['val']['left']['color'] = $color;
                            $result['tag']['left'] .= '><color rgb="' . $color . '"/></left>';
                        }
                        else {
                            $result['tag']['left'] .= '/>';
                        }
                    }
                }

                if ($side & self::BORDER_RIGHT) {
                    $style = isset($sideOptions['style']) ? self::_borderStyleName($sideOptions['style']) : null;
                    $color = isset($sideOptions['color']) ? self::normalizeColor($sideOptions['color']) : null;
                    if ($style || $color) {
                        if ($style) {
                            $result['val']['right']['style'] = $style;
                        }
                        $result['tag']['right'] = '<right' . ($style ? ' style="' . $style . '"' : '');
                        if ($color) {
                            $result['val']['right']['color'] = $color;
                            $result['tag']['right'] .= '><color rgb="' . $color . '"/></right>';
                        }
                        else {
                            $result['tag']['right'] .= '/>';
                        }
                    }
                }

                if ($side & self::BORDER_TOP) {
                    $style = isset($sideOptions['style']) ? self::_borderStyleName($sideOptions['style']) : null;
                    $color = isset($sideOptions['color']) ? self::normalizeColor($sideOptions['color']) : null;
                    if ($style || $color) {
                        if ($style) {
                            $result['val']['top']['style'] = $style;
                        }
                        $result['tag']['top'] = '<top' . ($style ? ' style="' . $style . '"' : '');
                        if ($color) {
                            $result['val']['top']['color'] = $color;
                            $result['tag']['top'] .= '><color rgb="' . $color . '"/></top>';
                        }
                        else {
                            $result['tag']['top'] .= '/>';
                        }
                    }
                }

                if ($side & self::BORDER_BOTTOM) {
                    $style = isset($sideOptions['style']) ? self::_borderStyleName($sideOptions['style']) : null;
                    $color = isset($sideOptions['color']) ? self::normalizeColor($sideOptions['color']) : null;
                    if ($style || $color) {
                        if ($style) {
                            $result['val']['bottom']['style'] = $style;
                        }
                        $result['tag']['bottom'] = '<bottom' . ($style ? ' style="' . $style . '"' : '');
                        if ($color) {
                            $result['val']['bottom']['color'] = $color;
                            $result['tag']['bottom'] .= '><color rgb="' . $color . '"/></bottom>';
                        }
                        else {
                            $result['tag']['bottom'] .= '/>';
                        }
                    }
                }

            }
        }

        $result['tag'] = implode('', $result['tag']);

        return $result;
    }

    /**
     * @param array|string $fill
     *
     * @return array
     */
    public static function normalizeFill($fill): array
    {
        $result = [];

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

            if (!empty($fillColor)) {
                if (!empty($fill['pattern'])) {
                    $result['val']['pattern'] = $fill['pattern'];
                    $result['tag']['patternFill'] = '<patternFill patternType="' . $fill['pattern'] . '">';
                }
                else {
                    $result['val']['pattern'] = 'solid';
                    $result['tag']['patternFill'] = '<patternFill patternType="solid">';
                }
                $result['tag']['fgColor'] = '<fgColor rgb="' . $fillColor . '"/>';
                $result['tag']['bgColor'] = '<bgColor indexed="64"/>';
                $result['tag']['/patternFill'] = '</patternFill>';
            }
            else {
                if (!empty($fill['pattern'])) {
                    $result['val']['pattern'] = $fill['pattern'];
                    $result['tag']['patternFill'] = '<patternFill patternType="' . $fill['pattern'] . '"/>';
                }
                else {
                    $result['val']['pattern'] = 'none';
                    $result['tag']['patternFill'] = '<patternFill patternType="none"/>';
                }
            }
        }
        $result['tag'] = '<fill>' . implode('', $result['tag']) . '</fill>';

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
        $defaultFontsNames = [
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

        foreach ($defaultFontsNames as $name => $defFont) {
            if (strcasecmp($fontName, $name) === 0) {
                return [$defFont['name'], $defFont['family']];
            }
        }
        return [null, null];
    }

    /**
     * @param array|string $font
     * @param string|null $key
     *
     * @return array
     */
    public static function normalizeFont($font, string &$key = null): array
    {
        $result = self::$instance->defaultFont;
        $result['tag'] = [];
        if (is_string($font)) {
            if (in_array($font, self::$fontStyleDefines, true)) {
                $font = ['style' => $font];
            }
            else {
                $font = [];
            }
        }
        if (!empty($font['name'])) {
            [$fontName, $fontFamily] = self::_getFamilyFont($font['name']);
            $result['val']['name'] = $fontName;
            $result['tag']['name'] = '<name val="' . $fontName . '"/><charset val="1"/><family val="' . (int)$fontFamily . '"/>';
        }
        if (!empty($font['size'])) {
            $size = (int)$font['size'];
            if ($size > 0) {
                $result['val']['name'] = $size;
                $result['tag']['size'] = '<sz val="' . $size . '"/>';
            }
        }
        if (!empty($font['color'])) {
            $color = self::normalizeColor($font['color']);
            if ($color) {
                $result['val']['name'] = $color;
                $result['tag']['color'] = '<color rgb="' . $color . '"/>';
            }
        }
        if (!empty($font['style'])) {
            if (is_array($font['style'])) {
                $val = implode('-', $font['style']);
            }
            else {
                $val = (string)$font['style'];
            }
            if ($val) {
                if (strpos($val, 'bold') !== false) {
                    $result['val']['style-bold'] = 1;
                    $result['tag']['style-bold'] = '<b/>';
                }
                if (strpos($val, 'italic') !== false) {
                    $result['val']['style-italic'] = 1;
                    $result['tag']['style-italic'] = '<i/>';
                }
                if (strpos($val, 'strike') !== false) {
                    $result['val']['style-strike'] = 1;
                    $result['tag']['style-strike'] = '<strike/>';
                }
                if (strpos($val, 'underline') !== false) {
                    $result['val']['style-underline'] = 1;
                    $result['tag']['style-underline'] = '<u/>';
                }
            }
        }
        $result['tag'] = '<font>' . implode('', $result['tag']) . '</font>';

        return $result;
    }

    /**
     * @param array $style
     *
     * @return array
     */
    public static function normalize(array $style): array
    {
        $result = [];
        foreach($style as $styleKey => $styleVal) {
            switch ($styleKey) {
                case 'format':
                    if ($styleVal === 0 || $styleVal === '0') {
                        $result[$styleKey] = '@INTEGER';
                    }
                    elseif ($styleVal && is_string($styleVal) && $styleVal[0] === '@') {
                        $result[$styleKey] = strtoupper($styleVal);
                    }
                    else {
                        $result[$styleKey] = $styleVal;
                    }
                    break;
                case 'border':
                    //$result[$styleKey] = self::normalizeBorder($styleVal);
                    $result[$styleKey] = self::borderOptions($styleVal);
                    break;
                case 'color':
                case 'text-color':
                case 'font-color':
                case 'fg-color':
                    $result['color'] = $styleVal;
                    break;
                case 'fill':
                case 'fill-color':
                case 'bg-color':
                case 'background-color':
                case 'cell-color':
                    $result['fill'] = $styleVal;
                    break;
                case 'font':
                    $result['font'] = $styleVal;
                    break;
                case 'align':
                    if ($styleVal === 'center' || $styleVal === 'center-center') {
                        $result['text-align'] = 'center';
                        $result['vertical-align'] = 'center';
                    }
                    elseif (strpos($styleVal, '-')) {
                        $parts = explode('-', $styleVal);
                        if (in_array($parts[0], ['general', 'left', 'right', 'justify'])) {
                            $result['text-align'] = $parts[0];
                            unset($parts[0]);
                        }
                        if (empty($result['text-align']) && in_array($parts[1], ['general', 'left', 'right', 'justify'])) {
                            $result['text-align'] = $parts[1];
                            unset($parts[1]);
                        }
                        if (!empty($parts[0]) && in_array($parts[0], ['bottom', 'center', 'distributed', 'top'])) {
                            $result['vertical-align'] = $parts[0];
                        }
                        if (!empty($parts[1]) && empty($result['vertical-align']) && in_array($parts[1], ['bottom', 'center', 'distributed', 'top'])) {
                            $result['vertical-align'] = $parts[1];
                            unset($parts[1]);
                        }
                    }
                    break;
                case 'text-align':
                case 'halign':
                case 'h-align':
                    if (in_array($styleVal, ['general', 'left', 'right', 'justify', 'center'])) {
                        $result['text-align'] = $styleVal;
                    }
                    break;
                case 'vertical-align':
                case 'valign':
                case 'v-align':
                    if (in_array($styleVal, ['bottom', 'center', 'distributed', 'top'])) {
                        $result['vertical-align'] = $styleVal;
                    }
                    break;
                case 'text-wrap':
                    $result['text-wrap'] = (bool)$styleVal;
                    break;
                case 'width':
                case 'autofit':
                    if ($styleVal === 'auto' || $styleVal === true) {
                        $result['options']['width-auto'] = true;
                    }
                    else {
                        $width = self::numFloat($styleVal);
                        if (is_numeric($width) && $width > 0) {
                            $result['width'] = $width;
                        }
                    }
                    break;
                default:
                    $result[$styleKey] = $styleVal;
            }
        }

        return $result;
    }

    /**
     * @param mixed $val
     *
     * @return mixed
     */
    public static function numFloat($val)
    {
        if (is_string($val)) {
            return (float)str_replace(',', '.', $val);
        }
        if (is_numeric($val)) {
            return (float)$val;
        }
        return $val;
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
     * @param string $sectionName
     * @param int $index
     *
     * @return array
     */
    protected function findElement(string $sectionName, int $index): array
    {
        if (isset($this->elementIndexes[$index], $this->elements[$sectionName][$this->elementIndexes[$index]])) {
            return $this->elements[$sectionName][$this->elementIndexes[$index]];
        }

        return [];
    }

    /**
     * @param string $sectionName
     * @param string|array $value
     * @param array|null $fullStyle
     *
     * @return int
     */
    protected function addElement(string $sectionName, $value, array $fullStyle = null): int
    {
        if (is_string($value)) {
            $key = $value;
        }
        elseif (isset($value['tag'])) {
            $key = $value['tag'];
        }
        else {
            $key = json_encode($value);
        }
        if (isset($this->elements[$sectionName][$key])) {
            return $this->elements[$sectionName][$key]['index'];
        }
        $index = isset($this->elements[$sectionName]) ? count($this->elements[$sectionName]) : 0;
        $this->elements[$sectionName][$key] = [
            'index' => $index,
            'value' => $value,
        ];
        if ($fullStyle) {
            $this->elements[$sectionName][$key]['style'] = $fullStyle;
        }
        $this->elementIndexes[$index] = $key;

        return $index;
    }

    /**
     * @param int $index
     *
     * @return array
     */
    protected function findStyleFont(int $index): array
    {
        return $this->findElement('fonts', $index);
    }

    /**
     * @param array $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    protected function addStyleFont(array &$cellStyle, array &$fullStyle = []): int
    {
        $index = 0;
        if (isset($cellStyle['font']) || isset($cellStyle['color']) || isset($cellStyle['text-color'])
            || isset($cellStyle['font-style']) || isset($cellStyle['font-size']))
        {
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

            if (!empty($cellStyle['font-underline'])) {
                if (is_numeric($cellStyle['font-underline']) && $cellStyle['font-underline'] >= 1 && $cellStyle['font-underline'] <= 2) {
                    $cellStyle['font']['style-underline'] = ($cellStyle['font-underline'] == 1) ? 'single' : 'double';
                }
                elseif (is_bool($cellStyle['font-underline'])) {
                    $cellStyle['font']['style-underline'] = 'single';
                }
                elseif (is_string($cellStyle['font-underline']) && in_array($cellStyle['font-underline'], ['single', 'double'])) {
                    $cellStyle['font']['style-underline'] = $cellStyle['font-underline'];
                }
                unset($cellStyle['font-underline']);
            }
            if (!empty($cellStyle['font-bold'])) {
                $cellStyle['font']['style-bold'] = 1;
                unset($cellStyle['font-bold']);
            }
            if (!empty($cellStyle['font-italic'])) {
                $cellStyle['font']['style-italic'] = 1;
                unset($cellStyle['font-italic']);
            }
            if (!empty($cellStyle['font-strike'])) {
                $cellStyle['font']['style-strike'] = 1;
                unset($cellStyle['font-strike']);
            }

            if (!empty($cellStyle['font-size']) && empty($cellStyle['font']['size'])) {
                $cellStyle['font']['size'] = $cellStyle['font-size'];
                unset($cellStyle['font-size']);
            }

            if ($cellStyle['font']) {
                $value = self::normalizeFont($cellStyle['font']);
                $index = $this->addElement('fonts', $value);
            }
            else {
                // optimization for default font
                $value = self::$instance->defaultFont;
                if (self::$instance->defaultFontId < 0) {
                    self::$instance->defaultFontId = $this->addElement('fonts', $value);
                }
                $index = self::$instance->defaultFontId;
            }

            if (isset($cellStyle['font'])) {
                unset($cellStyle['font']);
            }

            $fullStyle['font'] = $value;
        }
        else {
            $fullStyle['font'] = $this->findElement('fonts', $index);
        }
        $cellStyle['fontId'] = $index;

        return $index;
    }

    /**
     * @param int $index
     *
     * @return array
     */
    protected function findStyleFill(int $index): array
    {
        return $this->findElement('fills', $index);
    }

    /**
     * @param array $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    protected function addStyleFill(array &$cellStyle, array &$fullStyle = []): int
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

        return $index;
    }

    /**
     * @param int $index
     *
     * @return array
     */
    protected function findStyleBorder(int $index): array
    {
        return $this->findElement('borders', $index);
    }

    /**
     * @param array $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    protected function addStyleBorder(array &$cellStyle, array &$fullStyle = []): int
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

        return $index;
    }

    /**
     * @param array $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    protected function indexStyle(array $cellStyle, array &$fullStyle = []): int
    {
        if (isset($cellStyle['options'])) {
            $fullStyle['options'] = $cellStyle['options'];
            unset($cellStyle['options']);
        }
        //self::_ksort($cellStyle);

        return $this->addElement('cellXfs', $cellStyle, $fullStyle);
    }

    /**
     * @param string $format
     * @param array|null $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    public function addCellStyle(string $format, ?array $cellStyle = [], ?array &$fullStyle = []): int
    {
        $fullStyle = [];
        if (empty($cellStyle)) {
            $cellStyle = [];
        }
        if (!isset($cellStyle['numFmtId'])) {
            $cellStyle['numFmtId'] = 0;
        }
        if (!isset($cellStyle['fillId'])) {
            $cellStyle['fillId'] = 0;
        }
        if (!isset($cellStyle['borderId'])) {
            $cellStyle['borderId'] = 0;
        }
        if (!isset($cellStyle['xfId'])) {
            $cellStyle['xfId'] = 0;
        }
        $this->addStyleFont($cellStyle, $fullStyle);
        $this->addStyleFill($cellStyle, $fullStyle);
        $this->addStyleBorder($cellStyle, $fullStyle);

        if ($format) {
            $numberFormat = self::numberFormatStandardized($format, $xfId);
            $numberFormatType = self::determineNumberFormatType($numberFormat, $format);
            $cellStyle['numFmtId'] = $this->addElement('numFmts', $numberFormat);

            $fullStyle['format'] = $format;
            $fullStyle['number_format'] = $numberFormat;
            $fullStyle['number_format_type'] = $numberFormatType;
        }

        $cellXfsId = $this->indexStyle($cellStyle, $fullStyle);

        $fullStyle['cellXfsId'] = $cellXfsId;

        return $cellXfsId;
    }

    /**
     * @param array $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    public function addStyle(array $cellStyle, ?array &$fullStyle = []): int
    {
        if (isset($cellStyle['format'])) {
            $format = $cellStyle['format'];
            unset($cellStyle['format']);
        }
        else {
            $format = 'GENERAL';
        }

        return $this->addCellStyle($format, $cellStyle, $fullStyle);
    }

    /**
     * @param int $index
     *
     * @return array
     */
    public function findCellStyle(int $index): array
    {
        return $this->findElement('cellXfs', $index);
    }

    /**
     * @param string $sectionName
     *
     * @return array
     */
    protected function getElements(string $sectionName): array
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
    public function getStyleFonts(): array
    {
        return $this->getElements('fonts');
    }

    /**
     * @return array
     */
    public function getStyleFills(): array
    {
        return $this->getElements('fills');
    }

    /**
     * @return array
     */
    public function getStyleBorders(): array
    {
        return $this->getElements('borders');
    }

    /**
     * @return array
     */
    public function getStyleCellXfs(): array
    {
        return $this->getElements('cellXfs');
    }

    /**
     * @param string $numFormat
     * @param string|null $format
     *
     * @return string
     */
    private static function determineNumberFormatType(string $numFormat, string $format = null): string
    {
        if ($format === '@URL') {
            return 'n_shared_string';
        }
        if ($numFormat === 'GENERAL') {
            return 'n_auto';
        }
        if ($numFormat === '@') {
            return 'n_string';
        }
        if ($numFormat === '0') {
            return 'n_numeric';
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
        if (preg_match('/H{1,2}:M{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_datetime';
        }
        if (preg_match('/M{1,2}:S{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_datetime';
        }
        if (preg_match('/Y{2,4}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        if (preg_match('/D{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        if (preg_match('/M{1,2}(?![^"]*+")/i', $numFormat)) {
            return 'n_date';
        }
        return 'n_auto';
    }

    /**
     * @param $numFormat
     * @param int|null $xfId
     *
     * @return string
     */
    private static function numberFormatStandardized($numFormat, ?int &$xfId = 0): string
    {
        if (!$numFormat || !is_scalar($numFormat) || $numFormat === 'auto' || $numFormat === 'GENERAL') {
            return 'GENERAL';
        }
        if (is_int($numFormat)) {
            return '0';
        }
        if ($numFormat[0] === '@') {
            $numFormat = strtoupper($numFormat);
            if (strpos('@STRING', $numFormat) === 0 || strpos('@TEXT', $numFormat) === 0) {
                return '@';
            }
            if (strpos('@INTEGER', $numFormat) === 0) {
                return '0';
            }
            if (strpos('@PERCENT', $numFormat) === 0) {
                return '0%';
            }
        }

        while (isset(self::$instance->localeSettings['formats'][$numFormat])) {
            if (!$numFormat) {
                break;
            }
            if (isset(self::$instance->localeSettings['formats'][$numFormat])) {
                $numFormat = self::$instance->localeSettings['formats'][$numFormat];
            }
            else {
                break;
            }
        }

        $ignoreUntil = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($numFormat); $i < $ix; $i++) {
            $c = $numFormat[$i];

            if ($ignoreUntil === '' && $c === '[') {
                $ignoreUntil = ']';
            }
            elseif ($ignoreUntil === '' && $c === '"') {
                $ignoreUntil = '"';
            }
            elseif ($ignoreUntil === $c) {
                $ignoreUntil = '';
            }

            if ($ignoreUntil === '' && ($c === ' ' || $c === '-' || $c === '(' || $c === ')') && ($i === 0 || $numFormat[$i - 1] !== '_')) {
                $escaped .= "\\" . $c;
            }
            else {
                $escaped .= $c;
            }
        }

        return $escaped;
    }

    /**
     * @deprecated
     *
     * @param $format
     *
     * @return array
     */
    public function defineFormatType($format): array
    {
        static $defines = [];

        if (is_array($format)) {
            $format = reset($format);
        }

        if (!isset($defines[$format])) {
            $numberFormat = self::numberFormatStandardized($format);
            $numberFormatType = self::determineNumberFormatType($numberFormat);
            $cellStyleIdx = $this->addCellStyle($numberFormat, null);

            $defines[$format] = [
                'number_format' => $numberFormat, //contains excel format like 'YYYY-MM-DD HH:MM:SS'
                'number_format_type' => $numberFormatType, //contains friendly format like 'datetime'
                'default_style_idx' => $cellStyleIdx,
            ];
        }

        return $defines[$format];
    }

    /**
     * @return array
     */
    public function _getNumberFormats(): array
    {
        if (isset($this->elements['numFmts'])) {
            return array_keys($this->elements['numFmts']);
        }
        return [];
    }
}

// EOF