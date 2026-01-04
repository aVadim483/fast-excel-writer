# Class \avadim\FastExcelWriter\RichText\RichText

---

* [__construct()](#__construct) -- RichText constructor
* [__toString()](#__tostring)
* [addTaggedText()](#addtaggedtext) -- Add tagged text (<b>, <i>, <u>, <f>, <s>, <c>)
* [addText()](#addtext) -- Add a text fragment
* [setBold()](#setbold) -- Set bold font for the last added fragment
* [setColor()](#setcolor) -- Set font color for the last added fragment
* [setFont()](#setfont) -- Set font name for the last added fragment
* [fragment()](#fragment) -- Get fragment by its index
* [fragments()](#fragments) -- Get all fragments
* [setItalic()](#setitalic) -- Set italic font for the last added fragment
* [setSize()](#setsize) -- Set font size for the last added fragment
* [setUnderline()](#setunderline) -- Set underline for the last added fragment

---

## __construct()

---

```php
public function __construct($fragments)
```
_RichText constructor_

### Parameters

* `string|array|null $fragments`

---

## __toString()

---

```php
public function __toString(): string
```


### Parameters

_None_

---

## addTaggedText()

---

```php
public function addTaggedText(string $text): RichText
```
_Add tagged text (<b>, <i>, <u>, <f>, <s>, <c>)_

### Parameters

* `string $text`

---

## addText()

---

```php
public function addText(string $text, $prop): RichText
```
_Add a text fragment_

### Parameters

* `string $text`
* `mixed $prop`

---

## setBold()

---

```php
public function setBold(): RichText
```
_Set bold font for the last added fragment_

### Parameters

_None_

---

## setColor()

---

```php
public function setColor(string $color): RichText
```
_Set font color for the last added fragment_

### Parameters

* `string $color`

---

## setFont()

---

```php
public function setFont(string $font): RichText
```
_Set font name for the last added fragment_

### Parameters

* `string $font`

---

## fragment()

---

```php
public function fragment($num): RichTextFragment
```
_Get fragment by its index_

### Parameters

* `$num`

---

## fragments()

---

```php
public function fragments(): array
```
_Get all fragments_

### Parameters

_None_

---

## setItalic()

---

```php
public function setItalic(): RichText
```
_Set italic font for the last added fragment_

### Parameters

_None_

---

## setSize()

---

```php
public function setSize(int $size): RichText
```
_Set font size for the last added fragment_

### Parameters

* `int $size`

---

## setUnderline()

---

```php
public function setUnderline(): RichText
```
_Set underline for the last added fragment_

### Parameters

_None_

---

