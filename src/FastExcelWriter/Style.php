<?php

namespace avadim\FastExcelWriter;

class Style
{
    public const FORMAT = 'format';

    public const FONT = 'font';
    public const FONT_NAME = 'font-name';
    public const FONT_STYLE = 'font-style';
    public const FONT_STYLE_BOLD = 'font-style-bold';
    public const FONT_STYLE_ITALIC = 'font-style-italic';
    public const FONT_STYLE_UNDERLINE = 'font-style-underline';
    public const FONT_STYLE_STRIKE = 'font-style-strike';
    public const FONT_STYLE_STRIKETHROUGH = 'font-style-strike';

    public const FONT_SIZE = 'font-size';
    public const FONT_COLOR = 'font-color';

    public const STYLE = 'style';
    public const WIDTH = 'width';

    public const TEXT_WRAP = 'format-text-wrap';
    public const TEXT_ALIGN = 'format-align-horizontal';
    public const VERTICAL_ALIGN = 'format-align-vertical';

    public const TEXT_ALIGN_LEFT = 'left';
    public const TEXT_ALIGN_CENTER = 'center';
    public const TEXT_ALIGN_RIGHT = 'right';

    public const FILL_COLOR = 'fill-color';

    const FILL_SOLID = 'solid';
    const FILL_GRADIENT = 'linear';
    const FILL_GRADIENT_LINEAR = 'linear';

    public const FILL_PATTERN = 'fill-pattern';
    public const FILL_PATTERN_DARKDOWN = 'darkDown';
    public const FILL_PATTERN_DARKGRAY = 'darkGray';
    public const FILL_PATTERN_DARKGRID = 'darkGrid';
    public const FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal';
    public const FILL_PATTERN_DARKTRELLIS = 'darkTrellis';
    public const FILL_PATTERN_DARKUP = 'darkUp';
    public const FILL_PATTERN_DARKVERTICAL = 'darkVertical';
    public const FILL_PATTERN_GRAY0625 = 'gray0625';
    public const FILL_PATTERN_GRAY125 = 'gray125';
    public const FILL_PATTERN_LIGHTDOWN = 'lightDown';
    public const FILL_PATTERN_LIGHTGRAY = 'lightGray';
    public const FILL_PATTERN_LIGHTGRID = 'lightGrid';
    public const FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal';
    public const FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis';
    public const FILL_PATTERN_LIGHTUP = 'lightUp';
    public const FILL_PATTERN_LIGHTVERTICAL = 'lightVertical';
    public const FILL_PATTERN_MEDIUMGRAY = 'mediumGray';

    public const BORDER = 'border';

    public const BORDER_SIDE = 1;
    public const BORDER_STYLE = 'style';
    public const BORDER_COLOR = 'color';

    public const BORDER_LEFT = 1;
    public const BORDER_RIGHT = 2;
    public const BORDER_TOP = 4;
    public const BORDER_BOTTOM = 8;
    public const BORDER_ALL = self::BORDER_TOP + self::BORDER_RIGHT + self::BORDER_BOTTOM + self::BORDER_LEFT;

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


    protected array $styles = [];


    /**
     * @param array|null $style
     */
    public function __construct(?array $style = null)
    {
        if ($style) {
            $this->styles = StyleManager::normalize($style);
        }
    }

    /**
     * @param string $primaryKey
     * @param array $options
     *
     * @return void
     */
    protected function _setStyleOptions(string $primaryKey, array $options)
    {
        foreach ($options as $key => $val) {
            $this->styles[$primaryKey][$key] = $val;
        }
    }
    
    
    /**
     * Sets all borders style
     *
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function setBorder(string $style, ?string $color = '#000000'): Style
    {
        $options = [
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

        $this->_setStyleOptions(Style::BORDER, $options, true);

        return $this;
    }

    /**
     * Styles and color for left border
     *
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function setBorderLeft(string $style, ?string $color = '#000000'): Style
    {
        $options = [
            'border-left-style' => $style,
            'border-left-color' => $color,
        ];

        $this->_setStyleOptions(Style::BORDER, $options);

        return $this;
    }

    /**
     * Styles and color for right border
     *
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function setBorderRight(string $style, ?string $color = '#000000'): Style
    {
        $options = [
            'border-right-style' => $style,
            'border-right-color' => $color,
        ];

        $this->_setStyleOptions(Style::BORDER, $options);

        return $this;
    }

    /**
     * Styles and color for top border
     *
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function setBorderTop(string $style, ?string $color = '#000000'): Style
    {
        $options = [
            'border-top-style' => $style,
            'border-top-color' => $color,
        ];

        $this->_setStyleOptions(Style::BORDER, $options);

        return $this;
    }

    /**
     * Styles and color for bottom border
     *
     * @param string $style
     * @param string|null $color
     *
     * @return $this
     */
    public function setBorderBottom(string $style, ?string $color = '#000000'): Style
    {
        $options = [
            'border-bottom-style' => $style,
            'border-bottom-color' => $color,
        ];

        $this->_setStyleOptions(Style::BORDER, $options);

        return $this;
    }

    /**
     * Font name, size, style and color
     *
     * @param string $fontName
     * @param int|null $fontSize
     * @param string|null $fontStyle
     * @param string|null $fontColor
     *
     * @return $this
     */
    public function setFont(string $fontName, ?int $fontSize = null, ?string $fontStyle = null, ?string $fontColor = null): Style
    {
        $font = ['font-name' => $fontName];
        if ($fontSize) {
            $font['font-size'] = $fontSize;
        }
        if ($fontStyle) {
            $font['font-style'] = $fontStyle;
        }
        if ($fontColor) {
            $font['font-color'] = $fontColor;
        }

        $this->_setStyleOptions('font', $font);

        return $this;
    }


    /**
     * Font name
     *
     * @param string $fontName
     *
     * @return $this
     */
    public function setFontName(string $fontName): Style
    {
        $this->_setStyleOptions('font', [Style::FONT_NAME => $fontName]);

        return $this;
    }

    /**
     * Font size
     *
     * @param float $fontSize
     *
     * @return $this
     */
    public function setFontSize(float $fontSize): Style
    {
        $this->_setStyleOptions('font', [Style::FONT_SIZE => $fontSize]);

        return $this;
    }

    /**
     * Font style
     *
     * @param string $fontStyle
     *
     * @return $this
     */
    public function setFontStyle(string $fontStyle): Style
    {
        $fontStyle = str_replace(['font-style-', 'font-'] , '', strtolower($fontStyle));
        $this->_setStyleOptions('font', ['font-style-' . $fontStyle => 1]);

        return $this;
    }

    /**
     * Bold font
     *
     * @return $this
     */
    public function setFontStyleBold(): Style
    {
        return $this->setFontStyle('bold');
    }

    /**
     * Italic font
     *
     * @return $this
     */
    public function setFontStyleItalic(): Style
    {
        return $this->setFontStyle('italic');
    }

    /**
     * Sets underline
     *
     * @param bool|null $double
     *
     * @return $this
     */
    public function setFontStyleUnderline(?bool $double = false): Style
    {
        $this->_setStyleOptions('font', [Style::FONT_STYLE_UNDERLINE => $double ? 2 : 1]);

        return $this;
    }

    /**
     * Sets strikethrough
     *
     * @return $this
     */
    public function setFontStyleStrikethrough(): Style
    {
        return $this->setFontStyle(Style::FONT_STYLE_STRIKETHROUGH);
    }

    /**
     * Font color
     *
     * @param string $color
     *
     * @return $this
     */
    public function setFontColor(string $color): Style
    {
        $this->_setStyleOptions('font', [Style::FONT_COLOR => $color]);

        return $this;
    }

    /**
     * Alias of 'setFontColor()'
     *
     * @param string $color
     *
     * @return $this
     */
    public function setColor(string $color): Style
    {
        return $this->setFontColor($color);
    }

    /**
     * Another alias of 'setFontColor()'
     *
     * @param string $color
     *
     * @return $this
     */
    public function setTextColor(string $color): Style
    {
        return $this->setFontColor($color);
    }

    /**
     * Fill background color
     *
     * @param string $color
     * @param string|null $pattern
     *
     * @return $this
     */
    public function setFillColor(string $color, ?string $pattern = null): Style
    {
        $this->_setStyleOptions('fill', [Style::FILL_COLOR => $color, Style::FILL_PATTERN => $pattern ?: 'solid']);

        return $this;
    }

    /**
     * Alias of 'setFillColor()'
     *
     * @param string $color
     * @param string|null $pattern
     *
     * @return $this
     */
    public function setBgColor(string $color, ?string $pattern = null): Style
    {
        return $this->setFillColor($color, $pattern);
    }

    /**
     * Fill background by gradient
     *
     * @param string $color1
     * @param string $color2
     * @param int|null $degree
     *
     * @return $this
     */
    public function setFillGradient(string $color1, string $color2, ?int $degree = null): Style
    {
        $this->_setStyleOptions('fill', [
            'fill-pattern' => Style::FILL_GRADIENT_LINEAR,
            'fill-gradient-start' => $color1,
            'fill-gradient-end' => $color2,
            'fill-gradient-degree' => $degree ?: 0,
        ]);

        return $this;
    }

    /**
     * Horizontal and vertical content align
     *
     * @param string $textAlign
     * @param string|null $verticalAlign
     *
     * @return $this
     */
    public function setTextAlign(string $textAlign, ?string $verticalAlign = null): Style
    {
        $options = ['format-align-horizontal' => $textAlign];
        if ($verticalAlign !== null) {
            $options['format-align-vertical'] = $verticalAlign;
        }
        $this->_setStyleOptions('format', $options);

        return $this;
    }

    /**
     * Vertical content align
     *
     * @param string $verticalAlign
     *
     * @return $this
     */
    public function setVerticalAlign(string $verticalAlign): Style
    {
        $this->_setStyleOptions('format', ['vertical-align' => $verticalAlign]);

        return $this;
    }

    /**
     * Center content by horizontal and vertical
     *
     * @return $this
     */
    public function setTextCenter(): Style
    {
        $this->_setStyleOptions('format', [Style::TEXT_ALIGN => 'center', Style::VERTICAL_ALIGN => 'center']);

        return $this;
    }

    /**
     * Apply left alignment to content
     *
     * @return $this
     */
    public function setAlignLeft(): Style
    {
        return $this->setTextAlign('left');
    }

    /**
     * Apply right alignment to content
     *
     * @return $this
     */
    public function setAlignRight(): Style
    {
        return $this->setTextAlign('right');
    }

    /**
     * Text wrap option
     *
     * @param bool|null $textWrap
     *
     * @return $this
     */
    public function setTextWrap(?bool $textWrap = true): Style
    {
        $this->_setStyleOptions('format', [Style::TEXT_WRAP => (int)$textWrap]);

        return $this;
    }

    /**
     * Text rotation
     *
     * @param int $degrees
     *
     * @return $this
     */
    public function setTextRotation(int $degrees): Style
    {
        $this->_setStyleOptions('format', [ 'format-text-rotation' => $degrees ] );

        return $this;
    }

    /**
     * Indent left
     *
     * @param int $indent
     *
     * @return $this
     */
    public function setIndentLeft(int $indent): Style
    {
        $options = ['format-align-horizontal' => 'left', 'format-align-indent' => $indent];
        $this->_setStyleOptions('format', $options);

        return $this;
    }

    /**
     * Indent right
     *
     * @param int $indent
     *
     * @return $this
     */
    public function setIndentRight(int $indent): Style
    {
        $options = ['format-align-horizontal' => 'right', 'format-align-indent' => $indent];
        $this->_setStyleOptions('format', $options);

        return $this;
    }

    /**
     * Indent distributed
     *
     * @param int $indent
     *
     * @return $this
     */
    public function setIndentDistributed(int $indent): Style
    {
        $options = ['format-align-horizontal' => 'distributed', 'format-align-indent' => $indent];
        $this->_setStyleOptions('format', $options);

        return $this;
    }

    /**
     * Sets format
     *
     * @param string|array $format
     *
     * @return $this
     */
    public function setFormat($format): Style
    {
        if (is_array($format)) {
            $this->_setStyleOptions('format', $format);
        }
        else {
            if ($format && $format[0] === '@') {
                $format = strtoupper($format);
            }
            $this->_setStyleOptions('format', ['format-pattern' => $format]);
        }

        return $this;
    }

    /**
     * Return style properties as array
     *
     * @return array
     */
    public function toArray(): array
    {
        return $this->styles;
    }
}