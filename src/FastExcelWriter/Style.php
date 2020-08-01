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

    public const STYLE      = 'style';
    public const WIDTH      = 'width';

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

    static protected $font;

    /**
     * @param string $styleName
     *
     * @return string|null
     */
    protected static function _borderStyleName($styleName)
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
     * @param $font
     */
    public static function setDefaultFont($font)
    {
        [$fontName, $fontFamily] = self::_getFamilyFont($font['name']);
        if ($fontFamily) {
            $font['name'] = $fontName;
            $font['family'] = $fontFamily;
        }
        self::$font = $font;
    }

    /**
     * 'thin' -> all sides are thin
     * ['top' => ['style' => 'thin']]
     * ['top' => ['style' => 'thin', 'color' => '#f00']]
     *
     * @param $border
     *
     * @return array
     */
    public static function normalizeBorder($border)
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
                            $resultOptions['color'] = self::normaliazeColor($sideOptions['color']);
                        }
                    } elseif ($sideOptions[0] === '#') {
                        $resultOptions['color'] = self::normaliazeColor($sideOptions);
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
        }
        return $result ?: null;
    }

    /**
     * @param $color
     *
     * @return string|null
     */
    public static function normaliazeColor($color)
    {
        if ($color) {
            if (strpos($color, '#') === 0) {
                $color = substr($color, 1, 6);
            }
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

    protected static function _getFamilyFont($fontName)
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

        foreach($defaultFonts as $name => $defFont) {
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
    public static function normalaizeFont($font)
    {
        $result = self::$font;

        if (!empty($font)) {
            if (is_string($font)) {
                if (in_array($font, ['bold', 'italic', 'strike', 'underline'], true)) {
                    $font = ['style' => $font];
                } else {
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
                    default:
                        $result[$key] = $val;
                }
            }
        }
        return $result;
    }

    /**
     * @param $style
     *
     * @return array
     */
    public static function normalize($style)
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
                        $result['color'] = $styleVal;
                        break;
                    case 'fill':
                    case 'bg-color':
                    case 'background-color':
                    case 'cell-color':
                        $result['fill'] = $styleVal;
                        break;
                    case 'font':
                        $result['font'] = self::normalaizeFont($styleVal);
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


}

// EOF