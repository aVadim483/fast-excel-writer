<?php

namespace avadim\FastExcelWriter;

class Style
{
    public const FORMAT = 'format';

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
}