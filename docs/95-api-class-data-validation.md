# Class \avadim\FastExcelWriter\DataValidation\DataValidation

---

* [__construct()](#__construct) -- DataValidation constructor
* [custom()](#custom) -- Make data validation as a custom rule
* [date()](#date) -- Make data validation as a date value
* [decimal()](#decimal) -- Make data validation as a decimal value
* [dropDown()](#dropdown) -- Make data validation as a dropdown list
* [expression()](#expression) -- Make data validation as an expression (alias of self::custom())
* [integer()](#integer) -- Make data validation as an integer value
* [list()](#list) -- Alias of dropDown()
* [make()](#make) -- Make a DataValidation instance
* [isNumber()](#isnumber) -- Checking if a cell value is a number
* [isText()](#istext) -- Checking if a cell value is a text
* [textLength()](#textlength) -- Make data validation as a text length
* [whole()](#whole) -- Alias of integer()
* [allowBlank()](#allowblank) -- Allow blank value
* [getAttributes()](#getattributes)
* [setError()](#seterror) -- Set error message (title and body)
* [setErrorMessage()](#seterrormessage) -- Error message body
* [setErrorStyle()](#seterrorstyle) -- Error style (action in case of error)
* [setErrorTitle()](#seterrortitle) -- Error message title
* [setFormula()](#setformula) -- Alias of setFormula1()
* [setFormula1()](#setformula1) -- Set formula 1 for data validation
* [setFormula2()](#setformula2) -- Set formula 2 for data validation
* [setOperator()](#setoperator)
* [setPrompt()](#setprompt) -- Set prompt
* [showDropDown()](#showdropdown) -- Show dropdown list
* [showErrorMessage()](#showerrormessage) -- Allow (or disallow) error message
* [showInputMessage()](#showinputmessage) -- Show input message
* [setSqref()](#setsqref)
* [toXml()](#toxml)

---

## __construct()

---

```php
public function __construct($type)
```
_DataValidation constructor_

### Parameters

* `$type`

---

## custom()

---

```php
public static function custom(string $formula): DataValidation
```
_Make data validation as a custom rule_

### Parameters

* `string $formula`

---

## date()

---

```php
public static function date(string $operator, $formulas): DataValidation
```
_Make data validation as a date value_

### Parameters

* `string $operator`
* `string|int|array $formulas`

---

## decimal()

---

```php
public static function decimal(string $operator, $formulas): DataValidation
```
_Make data validation as a decimal value_

### Parameters

* `string $operator`
* `string|int|array $formulas`

---

## dropDown()

---

```php
public static function dropDown($formulas): DataValidation
```
_Make data validation as a dropdown list_

### Parameters

* `array|string $formulas`

---

## expression()

---

```php
public static function expression(string $formula): DataValidation
```
_Make data validation as an expression (alias of self::custom())_

### Parameters

* `string $formula`

---

## integer()

---

```php
public static function integer(string $operator, $formulas): DataValidation
```
_Make data validation as an integer value_

### Parameters

* `string $operator`
* `string|int|array $formulas`

---

## list()

---

```php
public static function list($formulas): DataValidation
```
_Alias of dropDown()_

### Parameters

* `array|string $formulas`

---

## make()

---

```php
public static function make($type): DataValidation
```
_Make a DataValidation instance_

### Parameters

* `$type`

---

## isNumber()

---

```php
public static function isNumber(): DataValidation
```
_Checking if a cell value is a number_

### Parameters

_None_

---

## isText()

---

```php
public static function isText(): DataValidation
```
_Checking if a cell value is a text_

### Parameters

_None_

---

## textLength()

---

```php
public static function textLength(string $operator, $formulas): DataValidation
```
_Make data validation as a text length_

### Parameters

* `string $operator`
* `array|string $formulas`

---

## whole()

---

```php
public static function whole(string $operator, $formulas): DataValidation
```
_Alias of integer()_

### Parameters

* `string $operator`
* `string|int|array $formulas`

---

## allowBlank()

---

```php
public function allowBlank(?bool $allowBlank = true): DataValidation
```
_Allow blank value_

### Parameters

* `bool $allowBlank`

---

## getAttributes()

---

```php
public function getAttributes(): array
```


### Parameters

_None_

---

## setError()

---

```php
public function setError(string $errorMessage, 
                         ?string $errorTitle = null): DataValidation
```
_Set error message (title and body)_

### Parameters

* `string $errorMessage`
* `string|null $errorTitle`

---

## setErrorMessage()

---

```php
public function setErrorMessage(string $error): DataValidation
```
_Error message body_

### Parameters

* `string $error`

---

## setErrorStyle()

---

```php
public function setErrorStyle(string $errorStyle): DataValidation
```
_Error style (action in case of error)_

### Parameters

* `string $errorStyle`

---

## setErrorTitle()

---

```php
public function setErrorTitle(string $errorTitle): DataValidation
```
_Error message title_

### Parameters

* `string $errorTitle`

---

## setFormula()

---

```php
public function setFormula($formula): DataValidation
```
_Alias of setFormula1()_

### Parameters

* `int|float|string|array|null $formula`

---

## setFormula1()

---

```php
public function setFormula1($formula): DataValidation
```
_Set formula 1 for data validation_

### Parameters

* `int|float|string|array|null $formula`

---

## setFormula2()

---

```php
public function setFormula2($formula): DataValidation
```
_Set formula 2 for data validation_

### Parameters

* `int|float|string|array|null $formula`

---

## setOperator()

---

```php
public function setOperator(string $operator, $formula1, 
                            $formula2): DataValidation
```


### Parameters

* `string $operator`
* `$formula1`
* `$formula2`

---

## setPrompt()

---

```php
public function setPrompt(string $promptMessage, 
                          ?string $promptTitle = null): DataValidation
```
_Set prompt_

### Parameters

* `string $promptMessage`
* `string|null $promptTitle`

---

## showDropDown()

---

```php
public function showDropDown(?bool $showDropDown = true): DataValidation
```
_Show dropdown list_

### Parameters

* `bool $showDropDown`

---

## showErrorMessage()

---

```php
public function showErrorMessage(?bool $showErrorMessage = true): DataValidation
```
_Allow (or disallow) error message_

### Parameters

* `bool $showErrorMessage`

---

## showInputMessage()

---

```php
public function showInputMessage(?bool $showInputMessage = true): DataValidation
```
_Show input message_

### Parameters

* `bool|null $showInputMessage`

---

## setSqref()

---

```php
public function setSqref(avadim\FastExcelWriter\Sheet $sheet, 
                         string $sqref): DataValidation
```


### Parameters

* `Sheet $sheet`
* `string $sqref`

---

## toXml()

---

```php
public function toXml($formulaConverter): string
```


### Parameters

* `$formulaConverter`

---

