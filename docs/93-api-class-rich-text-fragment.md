# Class \avadim\FastExcelWriter\RichText\RichTextFragment

---

* [__construct()](#__construct) -- RichTextFragment constructor
* [setBold()](#setbold) -- Set font weight to bold
* [setColor()](#setcolor) -- Set font color
* [setFont()](#setfont) -- Set font name
* [setItalic()](#setitalic) -- Set font style to italic
* [setSize()](#setsize) -- Set font size
* [setStrike()](#setstrike) -- Set font decoration to strikethrough
* [getText()](#gettext) -- Get fragment text
* [setUnderline()](#setunderline) -- Set font decoration to underline

---

## __construct()

---

```php
public function __construct(?string $text = null, ?array $prop = null)
```
_RichTextFragment constructor_

### Parameters

* `string|null $text`
* `array|null $prop`

---

## setBold()

---

```php
public function setBold(): RichTextFragment
```
_Set font weight to bold_

### Parameters

_None_

---

## setColor()

---

```php
public function setColor(string $color): RichTextFragment
```
_Set font color_

### Parameters

* `string $color`

---

## setFont()

---

```php
public function setFont(string $font): RichTextFragment
```
_Set font name_

### Parameters

* `string $font`

---

## setItalic()

---

```php
public function setItalic(): RichTextFragment
```
_Set font style to italic_

### Parameters

_None_

---

## setSize()

---

```php
public function setSize(int $size): RichTextFragment
```
_Set font size_

### Parameters

* `int $size`

---

## setStrike()

---

```php
public function setStrike(): RichTextFragment
```
_Set font decoration to strikethrough_

### Parameters

_None_

---

## getText()

---

```php
public function getText(): string
```
_Get fragment text_

### Parameters

_None_

---

## setUnderline()

---

```php
public function setUnderline(?bool $double = false): RichTextFragment
```
_Set font decoration to underline_

### Parameters

* `bool|null $double`

---

