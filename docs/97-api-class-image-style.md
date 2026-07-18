# Class \avadim\FastExcelWriter\Style\ImageStyle

---

* [__construct()](#__construct) – ImageStyle constructor
* [height()](#height) – Set height of image
* [hyperlink()](#hyperlink) – Set URL of hyperlink
* [offset()](#offset) – Set offset in pixels relative to the left and top borders of the cell
* [set()](#set) – Set style options from array
* [toArray()](#toarray) – Return style options as array
* [width()](#width) – Set width of image

---

## __construct()

---

```php
public function __construct(array $options = [])
```
_ImageStyle constructor_

### Parameters

* `array $options`

---

## height()

---

```php
public function height($height): ImageStyle
```
_Set height of image_

### Parameters

* `int|float $height`

---

## hyperlink()

---

```php
public function hyperlink(string $hyperlink): ImageStyle
```
_Set URL of hyperlink_

### Parameters

* `string $hyperlink`

---

## offset()

---

```php
public function offset($x, $y): ImageStyle
```
_Set offset in pixels relative to the left and top borders of the cell_

### Parameters

* `int|float $x`
* `int|float $y`

---

## set()

---

```php
public function set(array $options): ImageStyle
```
_Set style options from array_

### Parameters

* `array $options`

---

## toArray()

---

```php
public function toArray(): array
```
_Return style options as array_

### Parameters

_None_

---

## width()

---

```php
public function width($width): ImageStyle
```
_Set width of image_

### Parameters

* `int|float $width`

---

