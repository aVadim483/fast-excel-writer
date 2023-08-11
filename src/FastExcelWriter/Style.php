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
    public const FONT_NAME          = 'font-name';
    public const FONT_STYLE         = 'font-style';
    public const FONT_STYLE_BOLD    = 'font-style-bold';
    public const FONT_STYLE_ITALIC  = 'font-style-italic';
    public const FONT_STYLE_UNDERLINE  = 'font-style-underline';
    public const FONT_STYLE_STRIKETHROUGH  = 'font-style-strikethrough';

    public const FONT_SIZE          = 'font-size';
    public const FONT_COLOR          = 'font-color';

    public const STYLE              = 'style';
    public const WIDTH              = 'width';

    public const TEXT_WRAP          = 'format-text-wrap';
    public const TEXT_ALIGN         = 'format-align-horizontal';
    public const VERTICAL_ALIGN     = 'format-align-vertical';

    public const TEXT_ALIGN_LEFT    = 'left';
    public const TEXT_ALIGN_CENTER  = 'center';
    public const TEXT_ALIGN_RIGHT   = 'right';

    public const FILL_COLOR         = 'fill-color';

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

    public const BORDER_STYLE_NONE = null;
    public const BORDER_STYLE_THIN = 'thin';
    public const BORDER_STYLE_MEDIUM = 'medium';
    public const BORDER_STYLE_THICK = 'thick';
    public const BORDER_STYLE_DASH_DOT = 'dashDot';
    public const BORDER_STYLE_DASH_DOT_DOT = 'dashDotDot';
    public const BORDER_STYLE_DASHED = 'dashed';
    public const BORDER_STYLE_DOTTED = 'dotted';
    public const BORDER_STYLE_DOUBLE = 'double';
    public const BORDER_STYLE_HAIR = 'hair';
    public const BORDER_STYLE_MEDIUM_DASH_DOT = 'mediumDashDot';
    public const BORDER_STYLE_MEDIUM_DASH_DOT_DOT = 'mediumDashDotDot';
    public const BORDER_STYLE_MEDIUM_DASHED = 'mediumDashed';
    public const BORDER_STYLE_SLANT_DASH_DOT = 'slantDashDot';

    public const BORDER_STYLE_MIN = self::BORDER_NONE;
    public const BORDER_STYLE_MAX = self::BORDER_SLANT_DASH_DOT;

    protected static $instance;

    public array $localeSettings = [];

    public array $defaultFont = [];
    protected int $defaultFontId = -1;

    public array $defaultStyle = [];

    /** @var array Specified styles for hyperlinks */
    public array $hyperlinkStyle = [];

    /** @var array Specified styles for formats '@...'  */
    public array $defaultFormatStyles = [];

    public array $indexedColors = ['00000000',
        '00FFFFFF', '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF', '00000000', '00FFFFFF',
        '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF', '00800000', '00008000', '00000080',
        '00808000', '00800080', '00008080', '00C0C0C0', '00808080', '009999FF', '00993366', '00FFFFCC', '00CCFFFF',
        '00660066', '00FF8080', '000066CC', '00CCCCFF', '00000080', '00FF00FF', '00FFFF00', '0000FFFF', '00800080',
        '00800000', '00008080', '000000FF', '0000CCFF', '00CCFFFF', '00CCFFCC', '00FFFF99', '0099CCFF', '00FF99CC',
        '00CC99FF', '00E3E3E3', '003366FF', '0033CCCC', '0099CC00', '00FFCC00', '00FF9900', '00FF6600', '00666699',
        '00969696', '00003366', '00339966', '00003300', '00333300', '00993300', '00993366', '00333399', '00333333',
    ];

    protected array $elements = [];

    protected array $elementIndexes = [];

    protected static array $fontStyleDefines = ['bold', 'italic', 'strike', 'underline'];

    protected static array $borderStyleDefines = [self::BORDER_STYLE_THIN, self::BORDER_STYLE_MEDIUM, self::BORDER_STYLE_THICK, self::BORDER_STYLE_DASH_DOT,
        self::BORDER_STYLE_DASH_DOT_DOT, self::BORDER_STYLE_DASHED, self::BORDER_STYLE_DOTTED, self::BORDER_STYLE_DOUBLE, self::BORDER_STYLE_HAIR,
        self::BORDER_STYLE_MEDIUM_DASH_DOT, self::BORDER_STYLE_MEDIUM_DASH_DOT_DOT, self::BORDER_STYLE_MEDIUM_DASHED, self::BORDER_STYLE_SLANT_DASH_DOT];



    public array $_styleCache = [];

    /**
     * Constructor of Style
     *
     * @param array|null $options
     */
    public function __construct(?array $options)
    {
        self::$instance = $this;
        $defaultFont = [
            'font-name' => 'Arial',
            'font-size' => 10,
        ];
        $defaultStyle = [];
        $defaultFormatStyles = [];
        $hyperlinkStyle = [
            'font' => ['font-color' => '0563C1', 'font-style-underline' => 1],
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
        $defaultFont = ['font' => $defaultFont];
        $this->addStyleFont($defaultFont);

        $this->setDefaultStyle($defaultStyle);

        $styleVal = [
            'val' => ['pattern' => 'gray125'],
            'tag' => '<fill><patternFill patternType="none"/></fill>',
        ];
        $index = $this->addElement('fills', $styleVal);
        $styleVal = [
            'val' => ['pattern' => 'gray125'],
            'tag' => '<fill><patternFill patternType="gray125"/></fill>',
        ];
        $index = $this->addElement('fills', $styleVal);

        $styleVal = [
            'val' => [],
            'tag' => '<border><left/><right/><top/><bottom/><diagonal/></border>',
        ];
        $index = $this->addElement('borders', $styleVal);

        $styleVal = [
            '_num_fmt_id' => 0,
            '_font_id' => 0,
            '_fill_id' => 0,
            '_border_id' => 0,
            '_xf_id' => 0,
        ];
        $this->addXfs($styleVal,);

        $this->hyperlinkStyle = $hyperlinkStyle;
        $this->defaultFormatStyles = $defaultFormatStyles;
    }

    /**
     * @param array $styles
     *
     * @return array
     */
    public static function mergeStyles(array $styles): array
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
    public function setDefaultFont(array $font): Style
    {
        $this->defaultFont = self::normalizeFont($font);

        return $this;
    }

    /**
     * @param array $style
     *
     * @return $this
     */
    public function setDefaultStyle(array $style): Style
    {
        $this->defaultStyle = $style;

        return $this;
    }

    /**
     * @param array $localeData
     *
     * @return $this
     */
    public function setLocaleSettings(array $localeData): Style
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
     * @param $border
     * @return array|string[]
     */
    public static function borderOptions($border): array
    {
        $result = [];
        if (is_string($border)) {
            if ($border[0] === '#') {
                // it's a color
                $result = [
                    'border-left-style' => 'thin',
                    'border-left-color' => $border,
                    'border-right-style' => 'thin',
                    'border-right-color' => $border,
                    'border-top-style' => 'thin',
                    'border-top-color' => $border,
                    'border-bottom-style' => 'thin',
                    'border-bottom-color' => $border,
                    'border-diagonal-up' => 0,
                    'border-diagonal-down' => 0,

                ];
            }
            else {
                // it's a style
                $result = [
                    'border-left-style' => $border,
                    'border-left-color' => '#000000',
                    'border-right-style' => $border,
                    'border-right-color' => '#000000',
                    'border-top-style' => $border,
                    'border-top-color' => '#000000',
                    'border-bottom-style' => $border,
                    'border-bottom-color' => '#000000',
                    'border-diagonal-up' => 0,
                    'border-diagonal-down' => 0,

                ];
            }
        }
        else {
            foreach($border as $side => $sideOptions) {
                if ($sideOptions === null) {
                    $style = 'none';
                }
                elseif (is_string($sideOptions) && in_array($sideOptions, self::$borderStyleDefines)) {
                    $style = $sideOptions;
                }
                else {
                    $style = $sideOptions['style'] ?? 'thin';
                }
                $color = $sideOptions['color'] ?? '#000000';
                if (!is_numeric($side)) {
                    switch (strtolower($side)) {
                        case 'all':
                        case 'style':
                        case 'border-style':
                        case 'border-style-all':
                            $result = [
                                'border-left-style' => $style,
                                'border-left-color' => $color,
                                'border-right-style' => $style,
                                'border-right-color' => $color,
                                'border-top-style' => $style,
                                'border-top-color' => $color,
                                'border-bottom-style' => $style,
                                'border-bottom-color' => $color,
                                'border-diagonal-up' => 0,
                                'border-diagonal-down' => 0,
                            ];
                            break;
                        case 'left':
                            $result = [
                                'border-left-style' => $style,
                                'border-left-color' => $color,
                            ];
                            break;
                        case 'right':
                            $result = [
                                'border-right-style' => $style,
                                'border-right-color' => $color,
                            ];
                            break;
                        case 'top':
                            $result = [
                                'border-top-style' => $style,
                                'border-top-color' => $color,
                            ];
                            break;
                        case 'bottom':
                            $result = [
                                'border-bottom-style' => $style,
                                'border-bottom-color' => $color,
                            ];
                            break;
                        default:
                            $result[$side] = $sideOptions;
                    }
                }
                else {
                    if ($side & self::BORDER_LEFT) {
                        $result['border-left-style'] = $style;
                        $result['border-left-color'] = $color;
                    }
                    if ($side & self::BORDER_RIGHT) {
                        $result['border-right-style'] = $style;
                        $result['border-right-color'] = $color;
                    }
                    if ($side & self::BORDER_TOP) {
                        $result['border-top-style'] = $style;
                        $result['border-top-color'] = $color;
                    }
                    if ($side & self::BORDER_BOTTOM) {
                        $result['border-bottom-style'] = $style;
                        $result['border-bottom-color'] = $color;
                    }
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
            'val' => [],
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
            if (!empty($border['border-left-style']) && $border['border-left-style'] !== 'none' && $border['border-left-color']) {
                $style = is_int($border['border-left-style']) ? self::_borderStyleName($border['border-left-style']) : $border['border-left-style'];
                $color = self::normalizeColor($border['border-left-color']);

                $result['val']['left-style'] = $style;
                $result['val']['left-color'] = $color;

                $result['tag']['left'] = '<left style="' . $style . '"><color rgb="' . $color . '"/></left>';
            }
            else {
                $result['tag']['left'] = '<left/>';
            }

            if (!empty($border['border-right-style']) && $border['border-right-style'] !== 'none' && $border['border-right-color']) {
                $style = is_int($border['border-right-style']) ? self::_borderStyleName($border['border-right-style']) : $border['border-right-style'];
                $color = self::normalizeColor($border['border-right-color']);

                $result['val']['right-style'] = $style;
                $result['val']['right-color'] = $color;

                $result['tag']['right'] = '<right style="' . $style . '"><color rgb="' . $color . '"/></right>';
            }
            else {
                $result['tag']['right'] = '<right/>';
            }

            if (!empty($border['border-top-style']) && $border['border-top-style'] !== 'none' && $border['border-top-color']) {
                $style = is_int($border['border-top-style']) ? self::_borderStyleName($border['border-top-style']) : $border['border-top-style'];
                $color = self::normalizeColor($border['border-top-color']);

                $result['val']['top-style'] = $style;
                $result['val']['top-color'] = $color;

                $result['tag']['top'] = '<top style="' . $style . '"><color rgb="' . $color . '"/></top>';
            }
            else {
                $result['tag']['top'] = '<top/>';
            }

            if (!empty($border['border-bottom-style']) && $border['border-bottom-style'] !== 'none' && $border['border-bottom-color']) {
                $style = is_int($border['border-bottom-style']) ? self::_borderStyleName($border['border-bottom-style']) : $border['border-bottom-style'];
                $color = self::normalizeColor($border['border-bottom-color']);

                $result['val']['bottom-style'] = $style;
                $result['val']['bottom-color'] = $color;

                $result['tag']['bottom'] = '<bottom style="' . $style . '"><color rgb="' . $color . '"/></bottom>';
            }
            else {
                $result['tag']['bottom'] = '<bottom/>';
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
            if (!empty($fill['fill-color'])) {
                $fillColor = self::normalizeColor($fill['fill-color']);
            }
            elseif (!empty($fill['color']) ) {
                $fillColor = self::normalizeColor($fill['color']);
            }
            elseif (!empty($fill['fill']) && is_string($fill['fill']) && $fill['fill'] !== 'none') {
                $fillColor = self::normalizeColor($fill['fill']);
            }
            elseif (!empty($fill['bg-color']) && $fill['bg-color'] !== 'none') {
                $fillColor = self::normalizeColor($fill['bg-color']);
            }
            elseif (!empty($fill['background-color']) && $fill['background-color'] !== 'none') {
                $fillColor = self::normalizeColor($fill['background-color']);
            }

        }
        if (!empty($fillColor)) {
            if (!empty($fill['fill-pattern'])) {
                $fillPattern = $fill['fill-pattern'];
            }
            elseif (!empty($fill['pattern']) ) {
                $fillPattern = $fill['pattern'];
            }
            else {
                $fillPattern = 'solid';
            }
            $result['val']['color'] = $fillColor;
            $result['val']['pattern'] = $fillPattern;
            $result['tag'][] = '<patternFill patternType="' . $fillPattern . '">';
            $result['tag'][] = '<fgColor rgb="' . $fillColor . '"/>';
            $result['tag'][] = '<bgColor indexed="64"/>';
            $result['tag'][] = '</patternFill>';
        }
        else {
            $result['tag'][] = '<patternFill patternType="none"/>';
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
        static $normColors = [];

        if ($color) {
            if (isset($normColors[$color])) {
                return $normColors[$color];
            }

            if (strpos($color, '#') === 0) {
                $resColor = strtoupper(substr($color, 1, 6));
            }
            else {
                $resColor = strtoupper($color);
            }
            if (preg_match('/^[0-9A-F]+$/i', $resColor)) {
                if (strlen($resColor) === 3) {
                    $resColor = $resColor[0] . $resColor[0] . $resColor[1] . $resColor[1] . $resColor[2] . $resColor[2];
                }
                if (strlen($resColor) === 6) {
                    $resColor = 'FF' . $resColor;
                }
                if (strlen($resColor) > 8) {
                    $normColors[$color] = substr($resColor, 1, 8);
                }
                elseif (strlen($resColor) === 8) {
                    $normColors[$color] = $resColor;
                }
                else {
                    $normColors[$color] = null;
                }
            }
        }
        return $normColors[$color] ?? null;
    }

    /**
     * @param $fontName
     *
     * @return array|null[]
     */
    protected static function _getFamilyFont($fontName): array
    {
        static $defaultFontsNames = [
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
                return [$defFont['name'], $defFont['family'], 1];
            }
        }

        $defaultFontsNames[$fontName] = [
            'name' => $fontName,
            'family' => 0,
        ];

        return [$fontName, $defaultFontsNames[$fontName]['family'], 1];
    }

    /**
     * @param array|string $font
     *
     * @return array
     */
    public static function normalizeFont($font): array
    {
        $result = self::$instance->defaultFont;
        $result['tag'] = [];
        if (is_string($font)) {
            if (in_array($font, self::$fontStyleDefines, true)) {
                $font = ['font-style' => $font];
            }
            else {
                $font = ['font-name' => $font];
            }
        }

        $name = $font['font-name'] ?? ($font['name'] ?? null);
        if ($name) {
            [$fontName, $fontFamily, $fontCharset] = self::_getFamilyFont($name);
            if (isset($font['font-family'])) {
                $fontFamily = $font['font-family'];
            }
            if (isset($font['font-charset'])) {
                $fontCharset = $font['font-charset'];
            }
            $result['val']['name'] = $fontName;
            $result['tag']['name'] = '<name val="' . $fontName . '"/><charset val="' . $fontCharset . '"/><family val="' . (int)$fontFamily . '"/>';
            $result['font']['font-name'] = $fontName;
            $result['font']['font-family'] = $fontFamily;
            $result['font']['font-charset'] = $fontCharset;
        }

        $size = $font['font-size'] ?? ($font['size'] ?? null);
        if ($size) {
            $size = (int)$size;
            if ($size > 0) {
                $result['val']['size'] = $size;
                $result['tag']['size'] = '<sz val="' . $size . '"/>';
                $result['font']['font-size'] = $size;
            }
        }

        $color = $font['font-color'] ?? ($font['color'] ?? null);
        if ($color) {
            $color = self::normalizeColor($color);
            if ($color) {
                $result['val']['color'] = $color;
                $result['tag']['color'] = '<color rgb="' . $color . '"/>';
                $result['font']['font-color'] = $color;
            }
        }

        $style = $font['font-style'] ?? ($font['style'] ?? null);
        if ($style) {
            if (is_array($style)) {
                $val = implode('-', $style);
            }
            else {
                $val = (string)$style;
            }
            if ($val) {
                // backward compatibility
                if (strpos($val, 'bold') !== false) {
                    $result['val']['style-bold'] = 1;
                    $result['tag']['style-bold'] = '<b/>';
                    $result['font']['font-style-bold'] = 1;
                }
                if (strpos($val, 'italic') !== false) {
                    $result['val']['style-italic'] = 1;
                    $result['tag']['style-italic'] = '<i/>';
                    $result['font']['font-style-italic'] = 1;
                }
                if (strpos($val, 'strike') !== false) {
                    $result['val']['style-strike'] = 1;
                    $result['tag']['style-strike'] = '<strike/>';
                    $result['font']['font-style-strike'] = 1;
                }
                if (strpos($val, 'underline') !== false) {
                    $result['val']['style-underline'] = 1;
                    $result['tag']['style-underline'] = '<u/>';
                    $result['font']['font-style-underline'] = 1;
                }
            }
        }

        if (!empty($font['font-style-bold'])) {
            $result['val']['style-bold'] = 1;
            $result['tag']['style-bold'] = '<b/>';
            $result['font']['font-style-bold'] = 1;
        }
        if (!empty($font['font-style-italic'])) {
            $result['val']['style-italic'] = 1;
            $result['tag']['style-italic'] = '<i/>';
            $result['font']['font-style-italic'] = 1;
        }
        if (!empty($font['font-style-strike'])) {
            $result['val']['style-strike'] = 1;
            $result['tag']['style-strike'] = '<strike/>';
            $result['font']['font-style-strike'] = 1;
        }
        if (!empty($font['font-style-underline'])) {
            if ($font['font-style-underline'] === 'double' || (int)$font['font-style-underline'] === 2) {
                $result['val']['style-underline'] = 2;
                $result['tag']['style-underline'] = '<u val="double"/>';
                $result['font']['font-style-underline'] = 2;
            }
            else {
                $result['val']['style-underline'] = 1;
                $result['tag']['style-underline'] = '<u/>';
                $result['font']['font-style-underline'] = 1;
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
                        $result['format']['format-pattern'] = '@INTEGER';
                    }
                    elseif ($styleVal) {
                        if (is_array($styleVal)) {
                            $result['format'] = $styleVal;
                        }
                        elseif (is_string($styleVal) && $styleVal[0] === '@') {
                            $result['format']['format-pattern'] = strtoupper($styleVal);
                        }
                        else {
                            $result['format']['format-pattern'] = $styleVal;
                        }
                    }
                    break;

                case 'border':
                case 'border-style':
                case 'border-color':
                    $border = self::borderOptions($styleVal);
                    if (isset($result['border'])) {
                        $result['border'] = array_replace($result['border'], $border);
                    }
                    else {
                        $result['border'] = $border;
                    }
                    break;

                case 'font':
                    if (is_string($styleVal)) {
                        if (in_array($styleVal, self::$fontStyleDefines)) {
                            $result['font']['font-style'] = $styleVal;
                        }
                        else {
                            $result['font']['font-name'] = $styleVal;
                        }
                    }
                    elseif (is_array($styleVal)) {
                        $result['font'] = $styleVal;
                    }
                    break;

                case 'font-name':
                    $result['font']['font-name'] = $styleVal;
                    break;

                case 'font-color':
                case 'color':
                case 'text-color':
                case 'fg-color':
                    $result['font']['font-color'] = $styleVal;
                    break;

                case 'font-size':
                    $result['font']['font-size'] = $styleVal;
                    break;

                case 'fill':
                    if (is_array($styleVal)) {
                        $result['fill'] = $styleVal;
                    }
                    else {
                        $result['fill']['fill-color'] = $styleVal;
                        $result['fill']['fill-pattern'] = 'solid';
                    }
                    break;

                case 'fill-color':
                case 'bg-color':
                case 'background-color':
                    $result['fill']['fill-color'] = $styleVal;
                    if (!isset($result['fill']['fill-pattern'])) {
                        $result['fill']['fill-pattern'] = 'solid';
                    }
                    break;

                case 'fill-pattern':
                    $result['fill']['fill-pattern'] = $styleVal;
                    break;

                case 'text-rotation':
                    $result['format']['format-text-rotation'] = (int) $styleVal;
                    break;

                case 'align':
                case 'alignment':
                    if ($styleVal === 'center' || $styleVal === 'center-center') {
                        $result['format']['format-align-horizontal'] = 'center';
                        $result['format']['format-align-vertical'] = 'center';
                    }
                    elseif (strpos($styleVal, '-')) {
                        $parts = explode('-', $styleVal);
                        if (in_array($parts[0], ['general', 'left', 'right', 'justify'])) {
                            $result['format']['format-align-horizontal'] = $parts[0];
                            unset($parts[0]);
                        }
                        if (empty($result['format-align-horizontal']) && in_array($parts[1], ['general', 'left', 'right', 'justify'])) {
                            $result['format']['format-align-horizontal'] = $parts[1];
                            unset($parts[1]);
                        }
                        if (!empty($parts[0]) && in_array($parts[0], ['bottom', 'center', 'distributed', 'top'])) {
                            $result['format']['format-align-vertical'] = $parts[0];
                        }
                        if (!empty($parts[1]) && empty($result['vertical-align']) && in_array($parts[1], ['bottom', 'center', 'distributed', 'top'])) {
                            $result['format']['format-align-vertical'] = $parts[1];
                            unset($parts[1]);
                        }
                    }
                    break;

                case 'text-align':
                case 'format-align-horizontal':
                case 'format-alignment-horizontal':
                    if (in_array($styleVal, ['general', 'left', 'right', 'justify', 'center'])) {
                        $result['format']['format-align-horizontal'] = $styleVal;
                    }
                    break;

                case 'vertical-align':
                case 'format-align-vertical':
                case 'format-alignment-vertical':
                    if (in_array($styleVal, ['bottom', 'center', 'distributed', 'top'])) {
                        $result['format']['format-align-vertical'] = $styleVal;
                    }
                    break;

                case 'text-wrap':
                case 'format-text-wrap':
                    $result['format']['format-text-wrap'] = (bool)$styleVal;
                    break;

                case 'width':
                case 'col-width':
                    if ($styleVal === 'auto' || $styleVal === true) {
                        $result['options']['width-auto'] = true;
                    }
                    else {
                        $width = self::numFloat($styleVal);
                        if (is_numeric($width) && $width > 0) {
                            $result['options']['width'] = $width;
                        }
                    }
                    break;

                case 'height':
                case 'row-height':
                    $height = self::numFloat($styleVal);
                    if (is_numeric($height) && $height > 0) {
                        $result['options']['height'] = $height;
                    }
                    break;

                default:
                    if ($styleKey === 'font-style') {
                        $result['font']['font-style-' . $styleVal] = 1;
                    }
                    elseif (strpos($styleKey, 'font-') === 0) {
                        $result['font'][$styleKey] = $styleVal;
                    }
                    else {
                        $result[$styleKey] = $styleVal;
                    }
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
        if (isset($this->elementIndexes[$sectionName][$index], $this->elements[$sectionName][$this->elementIndexes[$sectionName][$index]])) {
            return $this->elements[$sectionName][$this->elementIndexes[$sectionName][$index]];
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
        $this->elementIndexes[$sectionName][$index] = $key;

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
        if (!empty($cellStyle['font'])) {
            if (is_string($cellStyle['font'])) {
                if (in_array($cellStyle['font'], self::$fontStyleDefines, true)) {
                    $cellStyle['font'] = ['font-style-' . $cellStyle['font'] => 1];
                }
                else {
                    $cellStyle['font'] = [];
                }
            }
        }
        else {
            $cellStyle['font'] = [];
        }

        if (!empty($cellStyle['font-name'])) {
            $cellStyle['font']['font-name'] = $cellStyle['font-name'];
            unset($cellStyle['font-name']);
        }
        if (!empty($cellStyle['font-size'])) {
            $cellStyle['font']['font-size'] = $cellStyle['font-size'];
            unset($cellStyle['font-size']);
        }
        if (!empty($cellStyle['font-color'])) {
            $cellStyle['font']['font-color'] = $cellStyle['font-color'];
            unset($cellStyle['font-color']);
        }

        if (!empty($cellStyle['font'])) {
            if (empty($cellStyle['font']['font-size']) && !empty($cellStyle['font']['size'])) {
                $cellStyle['font']['font-size'] = $cellStyle['font']['size'];
                unset($cellStyle['font']['size']);
            }
            if (empty($cellStyle['font']['font-color'])) {
                if (!empty($cellStyle['color'])) {
                    $cellStyle['font']['font-color'] = $cellStyle['color'];
                    unset($cellStyle['color']);
                }
                elseif (!empty($cellStyle['text-color'])) {
                    $cellStyle['font']['font-color'] = $cellStyle['text-color'];
                    unset($cellStyle['text-color']);
                }
            }

            if (!empty($cellStyle['font-style'])) {
                $cellStyle['font']['font-style-' . $cellStyle['font-style']] = 1;
                unset($cellStyle['font-style']);
            }
            elseif (!empty($cellStyle['font']['style'])) {
                $cellStyle['font']['font-style-' . $cellStyle['font']['style']] = 1;
                unset($cellStyle['font']['style']);
            }

            if (!empty($cellStyle['font-underline'])) {
                if ($cellStyle['font-underline'] === 'single') {
                    $cellStyle['font']['font-style-underline'] = 1;
                }
                elseif ($cellStyle['font-underline'] === 'double') {
                    $cellStyle['font']['font-style-underline'] = 2;
                }
                else {
                    $cellStyle['font']['font-style-underline'] = (int)$cellStyle['font-underline'];
                }
                unset($cellStyle['font-underline']);
            }
            if (!empty($cellStyle['font-bold'])) {
                $cellStyle['font']['font-style-bold'] = 1;
                unset($cellStyle['font-bold']);
            }
            if (!empty($cellStyle['font-italic'])) {
                $cellStyle['font']['font-style-italic'] = 1;
                unset($cellStyle['font-italic']);
            }
            if (!empty($cellStyle['font-strike'])) {
                $cellStyle['font']['font-style-strike'] = 1;
                unset($cellStyle['font-strike']);
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
        $cellStyle['_font_id'] = $index;

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
                $fill['fill-color'] = $cellStyle['fill'];
            }
            unset($cellStyle['fill']);
        }
        elseif (!empty($cellStyle['bg-color'])) {
            $fill['fill-color'] = $cellStyle['bg-color'];
            unset($cellStyle['bg-color']);
        }
        elseif (!empty($cellStyle['background-color'])) {
            $fill['fill-color'] = $cellStyle['background-color'];
            unset($cellStyle['background-color']);
        }

        if ($fill) {
            $value = self::normalizeFill($fill);
            $index = $this->addElement('fills', $value);

            $cellStyle['fill'] = $fill;
            $fullStyle['fills'] = $value;
        }
        if (!isset($fullStyle['fills'])) {
            $fullStyle['fills'] = $this->findElement('fills', $index);
        }
        $cellStyle['_fill_id'] = $index;

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
        if (!empty($cellStyle['border'])) {
            if (empty($this->elements['borders']) || $cellStyle['border'] !== 'none') {
                $value = self::normalizeBorder($cellStyle['border']);
                $index = $this->addElement('borders', $value);

                $fullStyle['borders'] = $value;
            }
            else {
                $fullStyle['borders'] = $this->findElement('borders', $index);
            }
            unset($cellStyle['border']);
        }
        $cellStyle['_border_id'] = $index;

        return $index;
    }

    /**
     * @param array $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    protected function addXfs(array $cellStyle, array &$fullStyle = []): int
    {
        if (isset($cellStyle['options'])) {
            $fullStyle['options'] = $cellStyle['options'];
            unset($cellStyle['options']);
        }

        return $this->addElement('cellXfs', $cellStyle, $fullStyle);
    }

    /**
     * @param string $numFormat
     * @param array|null $cellStyle
     * @param array|null $fullStyle
     *
     * @return int
     */
    public function addCellStyle(string $numFormat, ?array $cellStyle = [], ?array &$fullStyle = []): int
    {
        $fullStyle = [];
        if (empty($cellStyle)) {
            $cellStyle = [];
        }

        if (!isset($cellStyle['_fill_id'])) {
            $cellStyle['_fill_id'] = 0;
            if (isset($cellStyle['fill']) && $cellStyle['fill'] !== 'none') {
                $cellStyle['_fill_id'] = $this->addStyleFill($cellStyle, $fullStyle);
            }
        }
        if (!isset($cellStyle['_border_id'])) {
            $cellStyle['_border_id'] = 0;
            if (isset($cellStyle['border']) && $cellStyle['border'] !== 'none') {
                $cellStyle['_border_id'] = $this->addStyleBorder($cellStyle, $fullStyle);
            }
        }
        if (!isset($cellStyle['_font_id'])) {
            $cellStyle['_font_id'] = 0;
            if (isset($cellStyle['font']) && $cellStyle['font'] !== 'none') {
                $cellStyle['_font_id'] = $this->addStyleFont($cellStyle, $fullStyle);
            }
        }

        if (!isset($cellStyle['_xf_id'])) {
            $cellStyle['_xf_id'] = 0;
        }

        if ($numFormat && !isset($cellStyle['_num_fmt_id'])) {
            $cellStyle['_num_fmt_id'] = 0;

            $numberFormat = self::numberFormatStandardized($numFormat, $xfId);
            $numberFormatType = self::determineNumberFormatType($numberFormat, $numFormat);
            $cellStyle['_num_fmt_id'] = $this->addElement('numFmts', $numberFormat);

            $fullStyle['format']['format-pattern'] = $numFormat;
            $fullStyle['number_format'] = $numberFormat;
            $fullStyle['number_format_type'] = $numberFormatType;
        }
        else {
            $cellStyle['_num_fmt_id'] = 0;
        }

        $cellXfsId = $this->addXfs($cellStyle, $fullStyle);

        $fullStyle['_xf_id'] = $cellXfsId;

        return $cellXfsId;
    }

    /**
     * @param array $cellStyle
     * @param array|null $resultStyle
     *
     * @return int
     */
    public function addStyle(array $cellStyle, ?array &$resultStyle = []): int
    {
        if (isset($cellStyle['format']['format-pattern'])) {
            $numFormat = $cellStyle['format']['format-pattern'];
            unset($cellStyle['format']['format-pattern']);
        }
        else {
            $numFormat = 'GENERAL';
        }
        return $this->addCellStyle($numFormat, $cellStyle, $resultStyle);
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
     * @see https://support.microsoft.com/en-au/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68#ID0EDN
     *
     * #,###.00_);[Red](#,###.00);0.00;"gross receipts for "@
     * 1 - for positive numbers
     * 2 - for negative numbers
     * 3 - for zeros
     * 4 - for text
     *
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
            $numFormat = trim(strtoupper($numFormat));
            if (strpos('@STRING', $numFormat) === 0 || strpos('@TEXT', $numFormat) === 0) {
                return '@';
            }
            if (strpos('@INTEGER', $numFormat) === 0) {
                return '0';
            }
            if (strpos('@PERCENT', $numFormat) === 0) {
                return '0%';
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

            return $numFormat ?: '@';
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

            //if ($ignoreUntil === '' && ($c === ' ' || $c === '-' || $c === '(' || $c === ')') && ($i === 0 || $numFormat[$i - 1] !== '_')) {
            if ($ignoreUntil === '' && $c === ' ' && ($i === 0 || ($numFormat[$i - 1] !== '_' && $numFormat[$i - 1] !== '*'))) {
                $escaped .= "\\" . $c;
            }
            elseif ($ignoreUntil === '' && $c === '-' && ($i === 0 || $numFormat[$i - 1] === ']' || $numFormat[$i - 1] === ';')) {
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