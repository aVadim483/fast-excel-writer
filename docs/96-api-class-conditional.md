# Class \avadim\FastExcelWriter\Conditional\Conditional

---

* [__construct()](#__construct) -- Create a new Conditional
* [aboveAverage()](#aboveaverage)
* [beginsWith()](#beginswith) -- Applies a style if the cell value starts with the specified text
* [belowAverage()](#belowaverage)
* [between()](#between) -- The cell value is between two given values
* [colorScale()](#colorscale)
* [colorScaleMax()](#colorscalemax)
* [colorScaleMin()](#colorscalemin)
* [colorScaleNum()](#colorscalenum)
* [contains()](#contains) -- Applies a style if the cell value contains the specified text.
* [dataBar()](#databar) -- Colored data bar inside a cell
* [duplicateValues()](#duplicatevalues)
* [isEmpty()](#isempty) -- Applies a style if the cell is empty
* [endsWith()](#endswith) -- Applies a style if the cell value ends with the specified text
* [equals()](#equals) -- The cell value is equal to the given value
* [expression()](#expression) -- Applies the style if the expression evaluates to TRUE
* [greaterThan()](#greaterthan) -- The cell value is greater than the specified value
* [greaterThanOrEqual()](#greaterthanorequal) -- The cell value is greater than or equal to the specified value
* [lessThan()](#lessthan) -- The cell value is less than the specified value
* [lessThanOrEqual()](#lessthanorequal) -- The cell value is less than or equal to the specified value
* [low()](#low)
* [lowPercent()](#lowpercent)
* [make()](#make) -- Cell value is compared to a specified value or formula
* [notBetween()](#notbetween) -- The cell value is between two given values
* [notContains()](#notcontains) -- Applies a style if the cell value does not contain the specified text.
* [notEquals()](#notequals) -- The cell value is not equal to the specified value
* [top()](#top)
* [topPercent()](#toppercent)
* [uniqueValues()](#uniquevalues)
* [setDirectionRtl()](#setdirectionrtl) -- Determines the direction of the bars
* [setDxfId()](#setdxfid)
* [setFillColor()](#setfillcolor)
* [setFontColor()](#setfontcolor)
* [setGradient()](#setgradient) -- Enables or disables the gradient style of the bars
* [setShowValue()](#setshowvalue) -- Controls the display of the value in a cell
* [setSqref()](#setsqref)
* [getStyle()](#getstyle)
* [setStyle()](#setstyle)
* [toXml()](#toxml)

---

## __construct()

---

```php
public function __construct(string $type, string $operator, $options, $style)
```
_Create a new Conditional_

### Parameters

* `$type`
* `$operator`
* `$options`
* `$style`

---

## aboveAverage()

---

```php
public static function aboveAverage(array $style): Conditional
```


### Parameters

* `array $style`

---

## beginsWith()

---

```php
public static function beginsWith(string $text, 
                                  ?array $style = null): Conditional
```
_Applies a style if the cell value starts with the specified text_

### Parameters

* `string $text`
* `array|null $style`

---

## belowAverage()

---

```php
public static function belowAverage(array $style): Conditional
```


### Parameters

* `array $style`

---

## between()

---

```php
public static function between(array $values, 
                               ?array $style = null): Conditional
```
_The cell value is between two given values_

### Parameters

* `int[]|float[] $values`
* `array|null $style`

---

## colorScale()

---

```php
public static function colorScale(string $color1, string $color2, 
                                  ?string $color3 = null): Conditional
```


### Parameters

* `string $color1`
* `string $color2`
* `string|null $color3`

---

## colorScaleMax()

---

```php
public static function colorScaleMax(string $color): Conditional
```


### Parameters

* `string $color`

---

## colorScaleMin()

---

```php
public static function colorScaleMin(string $color): Conditional
```


### Parameters

* `string $color`

---

## colorScaleNum()

---

```php
public static function colorScaleNum(array $values, string $color1, 
                                     string $color2, 
                                     ?string $color3 = null): Conditional
```


### Parameters

* `array $values`
* `string $color1`
* `string $color2`
* `string|null $color3`

---

## contains()

---

```php
public static function contains(string $text, 
                                ?array $style = null): Conditional
```
_Applies a style if the cell value contains the specified text._

### Parameters

* `string $text`
* `array|null $style`

---

## dataBar()

---

```php
public static function dataBar(string $color): Conditional
```
_Colored data bar inside a cell_

### Parameters

* `string $color`

---

## duplicateValues()

---

```php
public static function duplicateValues(array $style): Conditional
```


### Parameters

* `array $style`

---

## isEmpty()

---

```php
public static function isEmpty(?string $cell = null, 
                               ?array $style = null): Conditional
```
_Applies a style if the cell is empty_

### Parameters

* `string|null $cell`
* `array|null $style`

---

## endsWith()

---

```php
public static function endsWith(string $text, 
                                ?array $style = null): Conditional
```
_Applies a style if the cell value ends with the specified text_

### Parameters

* `string $text`
* `array|null $style`

---

## equals()

---

```php
public static function equals($value, ?array $style = null): Conditional
```
_The cell value is equal to the given value_

### Parameters

* `int|float|string $value`
* `array|null $style`

---

## expression()

---

```php
public static function expression(string $formula, 
                                  ?array $style = null): Conditional
```
_Applies the style if the expression evaluates to TRUE_

### Parameters

* `string $formula`
* `array|null $style`

---

## greaterThan()

---

```php
public static function greaterThan($value, ?array $style = null): Conditional
```
_The cell value is greater than the specified value_

### Parameters

* `int|float|string $value`
* `array|null $style`

---

## greaterThanOrEqual()

---

```php
public static function greaterThanOrEqual($value, 
                                          ?array $style = null): Conditional
```
_The cell value is greater than or equal to the specified value_

### Parameters

* `int|float|string $value`
* `array|null $style`

---

## lessThan()

---

```php
public static function lessThan($value, ?array $style = null): Conditional
```
_The cell value is less than the specified value_

### Parameters

* `int|float|string $value`
* `array|null $style`

---

## lessThanOrEqual()

---

```php
public static function lessThanOrEqual($value, 
                                       ?array $style = null): Conditional
```
_The cell value is less than or equal to the specified value_

### Parameters

* `int|float|string $value`
* `array|null $style`

---

## low()

---

```php
public static function low(int $rank, array $style): Conditional
```


### Parameters

* `int $rank`
* `array $style`

---

## lowPercent()

---

```php
public static function lowPercent(int $rank, array $style): Conditional
```


### Parameters

* `int $rank`
* `array $style`

---

## make()

---

```php
public static function make(string $operator, $formula, 
                            ?array $style = null): Conditional
```
_Cell value is compared to a specified value or formula_

### Parameters

* `string $operator`
* `int|float|string|array $formula`
* `array|null $style`

---

## notBetween()

---

```php
public static function notBetween(array $values, 
                                  ?array $style = null): Conditional
```
_The cell value is between two given values_

### Parameters

* `int[]|float[] $values`
* `array|null $style`

---

## notContains()

---

```php
public static function notContains(string $text, 
                                   ?array $style = null): Conditional
```
_Applies a style if the cell value does not contain the specified text._

### Parameters

* `string $text`
* `array|null $style`

---

## notEquals()

---

```php
public static function notEquals($value, ?array $style = null): Conditional
```
_The cell value is not equal to the specified value_

### Parameters

* `int|float|string $value`
* `array|null $style`

---

## top()

---

```php
public static function top(int $rank, array $style): Conditional
```


### Parameters

* `int $rank`
* `array $style`

---

## topPercent()

---

```php
public static function topPercent(int $rank, array $style): Conditional
```


### Parameters

* `int $rank`
* `array $style`

---

## uniqueValues()

---

```php
public static function uniqueValues(array $style): Conditional
```


### Parameters

* `array $style`

---

## setDirectionRtl()

---

```php
public function setDirectionRtl(bool $value): Conditional
```
_Determines the direction of the bars_

### Parameters

* `bool $value`

---

## setDxfId()

---

```php
public function setDxfId(int $dxfId): Conditional
```


### Parameters

* `int $dxfId`

---

## setFillColor()

---

```php
public function setFillColor($color): Conditional
```


### Parameters

* `$color`

---

## setFontColor()

---

```php
public function setFontColor($color): Conditional
```


### Parameters

* `$color`

---

## setGradient()

---

```php
public function setGradient(bool $value): Conditional
```
_Enables or disables the gradient style of the bars_

### Parameters

* `bool $value`

---

## setShowValue()

---

```php
public function setShowValue(bool $value): Conditional
```
_Controls the display of the value in a cell_

### Parameters

* `bool $value`

---

## setSqref()

---

```php
public function setSqref(avadim\FastExcelWriter\Sheet $sheet, 
                         string $sqref): Conditional
```


### Parameters

* `Sheet $sheet`
* `string $sqref`

---

## getStyle()

---

```php
public function getStyle(): ?array
```


### Parameters

_None_

---

## setStyle()

---

```php
public function setStyle($style): Conditional
```


### Parameters

* `string|array $style`

---

## toXml()

---

```php
public function toXml(int $priority, $formulaConverter): string
```


### Parameters

* `int $priority`
* `$formulaConverter`

---

