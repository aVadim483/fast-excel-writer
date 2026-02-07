# Class \avadim\FastExcelWriter\Options

---

* [__construct()](#__construct) – Options constructor
* [create()](#create) – Create Options instance
* [autoConvertNumber()](#autoconvertnumber) – Set auto conversion for numbers
* [defaultFont()](#defaultfont) – Set default font
* [locale()](#locale) – Set locale
* [sharedString()](#sharedstring) – Use shared strings
* [styleManagerClass()](#stylemanagerclass) – Set Style Manager class
* [tempDir()](#tempdir) – Set temporary directory
* [tempPrefix()](#tempprefix) – Set prefix for temporary files
* [toArray()](#toarray) – Return options as array
* [writerClass()](#writerclass) – Set Writer class

---

## __construct()

---

```php
public function __construct(array $options = [])
```
_Options constructor_

### Parameters

* `array $options`

---

## create()

---

```php
public static function create(array $options = []): Options
```
_Create Options instance_

### Parameters

* `array $options`

---

## autoConvertNumber()

---

```php
public function autoConvertNumber(bool $autoConvertNumber = true): Options
```
_Set auto conversion for numbers_

### Parameters

* `bool $autoConvertNumber`

---

## defaultFont()

---

```php
public function defaultFont(array $fontOptions): Options
```
_Set default font_

### Parameters

* `array $fontOptions`

---

## locale()

---

```php
public function locale(string $locale): Options
```
_Set locale_

### Parameters

* `string $locale`

---

## sharedString()

---

```php
public function sharedString(bool $sharedString = true): Options
```
_Use shared strings_

### Parameters

* `bool $sharedString`

---

## styleManagerClass()

---

```php
public function styleManagerClass(string $styleManagerClass): Options
```
_Set Style Manager class_

### Parameters

* `string $styleManagerClass`

---

## tempDir()

---

```php
public function tempDir(string $tempDir): Options
```
_Set temporary directory_

### Parameters

* `string $tempDir`

---

## tempPrefix()

---

```php
public function tempPrefix(string $tempPrefix): Options
```
_Set prefix for temporary files_

### Parameters

* `string $tempPrefix`

---

## toArray()

---

```php
public function toArray(): array
```
_Return options as array_

### Parameters

_None_

---

## writerClass()

---

```php
public function writerClass(string $writerClass): Options
```
_Set Writer class_

### Parameters

* `string $writerClass`

---

