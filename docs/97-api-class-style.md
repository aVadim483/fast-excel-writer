# Class \avadim\FastExcelWriter\Style\Style

---

* [setAlignLeft()](#setalignleft) -- Apply left alignment to content
* [setAlignRight()](#setalignright) -- Apply right alignment to content
* [setBgColor()](#setbgcolor) -- Alias of 'setFillColor()'
* [setBorder()](#setborder) -- Sets all borders style
* [setBorderBottom()](#setborderbottom) -- Styles and color for bottom border
* [setBorderLeft()](#setborderleft) -- Styles and color for left border
* [setBorderRight()](#setborderright) -- Styles and color for right border
* [setBorderTop()](#setbordertop) -- Styles and color for top border
* [setColor()](#setcolor) -- Alias of 'setFontColor()'
* [setFillColor()](#setfillcolor) -- Fill background color
* [setFillGradient()](#setfillgradient) -- Fill background by gradient
* [setFont()](#setfont) -- Font name, size, style and color
* [setFontColor()](#setfontcolor) -- Font color
* [setFontName()](#setfontname) -- Font name
* [setFontSize()](#setfontsize) -- Font size
* [setFontStyle()](#setfontstyle) -- Font style
* [setFontStyleBold()](#setfontstylebold) -- Bold font
* [setFontStyleItalic()](#setfontstyleitalic) -- Italic font
* [setFontStyleStrikethrough()](#setfontstylestrikethrough) -- Sets strikethrough
* [setFontStyleUnderline()](#setfontstyleunderline) -- Sets underline
* [setFormat()](#setformat) -- Sets format
* [setIndentDistributed()](#setindentdistributed) -- Indent distributed
* [setIndentLeft()](#setindentleft) -- Indent left
* [setIndentRight()](#setindentright) -- Indent right
* [setTextAlign()](#settextalign) -- Horizontal and vertical content align
* [setTextCenter()](#settextcenter) -- Center content by horizontal and vertical
* [setTextColor()](#settextcolor) -- Another alias of 'setFontColor()'
* [setTextRotation()](#settextrotation) -- Text rotation
* [setTextWrap()](#settextwrap) -- Text wrap option
* [setVerticalAlign()](#setverticalalign) -- Vertical content align
* [toArray()](#toarray) -- Return style properties as array

---

## setAlignLeft()

---

```php
public function setAlignLeft(): Style
```
_Apply left alignment to content_

### Parameters

_None_

---

## setAlignRight()

---

```php
public function setAlignRight(): Style
```
_Apply right alignment to content_

### Parameters

_None_

---

## setBgColor()

---

```php
public function setBgColor(string $color, ?string $pattern = null): Style
```
_Alias of 'setFillColor()'_

### Parameters

* `string $color`
* `string|null $pattern`

---

## setBorder()

---

```php
public function setBorder(string $style, ?string $color = '#000000'): Style
```
_Sets all borders style_

### Parameters

* `string $style`
* `string|null $color`

---

## setBorderBottom()

---

```php
public function setBorderBottom(string $style, 
                                ?string $color = '#000000'): Style
```
_Styles and color for bottom border_

### Parameters

* `string $style`
* `string|null $color`

---

## setBorderLeft()

---

```php
public function setBorderLeft(string $style, 
                              ?string $color = '#000000'): Style
```
_Styles and color for left border_

### Parameters

* `string $style`
* `string|null $color`

---

## setBorderRight()

---

```php
public function setBorderRight(string $style, 
                               ?string $color = '#000000'): Style
```
_Styles and color for right border_

### Parameters

* `string $style`
* `string|null $color`

---

## setBorderTop()

---

```php
public function setBorderTop(string $style, ?string $color = '#000000'): Style
```
_Styles and color for top border_

### Parameters

* `string $style`
* `string|null $color`

---

## setColor()

---

```php
public function setColor(string $color): Style
```
_Alias of 'setFontColor()'_

### Parameters

* `string $color`

---

## setFillColor()

---

```php
public function setFillColor(string $color, ?string $pattern = null): Style
```
_Fill background color_

### Parameters

* `string $color`
* `string|null $pattern`

---

## setFillGradient()

---

```php
public function setFillGradient(string $color1, string $color2, 
                                ?int $degree = null): Style
```
_Fill background by gradient_

### Parameters

* `string $color1`
* `string $color2`
* `int|null $degree`

---

## setFont()

---

```php
public function setFont(string $fontName, ?int $fontSize = null, 
                        ?string $fontStyle = null, 
                        ?string $fontColor = null): Style
```
_Font name, size, style and color_

### Parameters

* `string $fontName`
* `int|null $fontSize`
* `string|null $fontStyle`
* `string|null $fontColor`

---

## setFontColor()

---

```php
public function setFontColor(string $color): Style
```
_Font color_

### Parameters

* `string $color`

---

## setFontName()

---

```php
public function setFontName(string $fontName): Style
```
_Font name_

### Parameters

* `string $fontName`

---

## setFontSize()

---

```php
public function setFontSize(float $fontSize): Style
```
_Font size_

### Parameters

* `float $fontSize`

---

## setFontStyle()

---

```php
public function setFontStyle(string $fontStyle): Style
```
_Font style_

### Parameters

* `string $fontStyle`

---

## setFontStyleBold()

---

```php
public function setFontStyleBold(): Style
```
_Bold font_

### Parameters

_None_

---

## setFontStyleItalic()

---

```php
public function setFontStyleItalic(): Style
```
_Italic font_

### Parameters

_None_

---

## setFontStyleStrikethrough()

---

```php
public function setFontStyleStrikethrough(): Style
```
_Sets strikethrough_

### Parameters

_None_

---

## setFontStyleUnderline()

---

```php
public function setFontStyleUnderline(?bool $double = false): Style
```
_Sets underline_

### Parameters

* `bool|null $double`

---

## setFormat()

---

```php
public function setFormat($format): Style
```
_Sets format_

### Parameters

* `string|array $format`

---

## setIndentDistributed()

---

```php
public function setIndentDistributed(int $indent): Style
```
_Indent distributed_

### Parameters

* `int $indent`

---

## setIndentLeft()

---

```php
public function setIndentLeft(int $indent): Style
```
_Indent left_

### Parameters

* `int $indent`

---

## setIndentRight()

---

```php
public function setIndentRight(int $indent): Style
```
_Indent right_

### Parameters

* `int $indent`

---

## setTextAlign()

---

```php
public function setTextAlign(string $textAlign, 
                             ?string $verticalAlign = null): Style
```
_Horizontal and vertical content align_

### Parameters

* `string $textAlign`
* `string|null $verticalAlign`

---

## setTextCenter()

---

```php
public function setTextCenter(): Style
```
_Center content by horizontal and vertical_

### Parameters

_None_

---

## setTextColor()

---

```php
public function setTextColor(string $color): Style
```
_Another alias of 'setFontColor()'_

### Parameters

* `string $color`

---

## setTextRotation()

---

```php
public function setTextRotation(int $degrees): Style
```
_Text rotation_

### Parameters

* `int $degrees`

---

## setTextWrap()

---

```php
public function setTextWrap(?bool $textWrap = true): Style
```
_Text wrap option_

### Parameters

* `bool|null $textWrap`

---

## setVerticalAlign()

---

```php
public function setVerticalAlign(string $verticalAlign): Style
```
_Vertical content align_

### Parameters

* `string $verticalAlign`

---

## toArray()

---

```php
public function toArray(): array
```
_Return style properties as array_

### Parameters

_None_

---

