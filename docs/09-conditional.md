## FastExcelWriter - Conditional Formatting (since v.6.4)

Conditional Formatting allows formatting options to be set based on the value of a cell.
It can be applied to individual cells, or to a range of cells.

### Simple usage

```php
use avadim\FastExcelWriter\Conditional\Conditional;
use avadim\FastExcelWriter\Style;

// You can set the foreground colour of a cell to dark red and background to light red if its value is below zero
$cond1 = Conditional::make('<', 0)->setStyle([Style::FONT_COLOR => '#900', Style::FILL_COLOR => '#f99']);
// You can set the foreground colour of a cell to dark green and background to light green if its value is more than 10
$cond2 = Conditional::make('>', 10)->setStyle([Style::FONT_COLOR => '#090', Style::FILL_COLOR => '#9f9']);

// Conditional formatting for single cell
$sheet->writeCell(11)
    ->applyConditionalFormatting([$cond1, $cond2]);

// Conditional formatting for column C
$sheet->addConditionalFormatting('C1:C' . Excel::MAX_ROW, [$cond1, $cond2]);
```

**IMPORTANT:** If you want to set conditional formatting for a large number of cells, it is not recommended 
to use the `applyConditionalFormatting()` function, as this can negatively affect performance. 
Use `$sheet->addConditionalFormatting($range, $conditional)` to set the same formatting conditions for a range of cells

### General cell value check

A base type in which the contents of a cell are compared to a given value.

```php
$value = 10;
$style = [Style::FONT_COLOR => '#900', Style::FILL_COLOR => '#f99'];

// Comparison with numeric values
$cond = Conditional::make('=', $value, $style);
$cond = Conditional::equals($value, $style); // the same result

$cond = Conditional::make('!=', $value, $style);
$cond = Conditional::make('<>', $value, $style);
$cond = Conditional::notEquals($value, $style);

$cond = Conditional::make('>', $value, $style);
$cond = Conditional::greaterThan($value, $style);

$cond = Conditional::make('>=', $value, $style);
$cond = Conditional::greaterThanOrEqual($value, $style);

$cond = Conditional::make('<', $value, $style);
$cond = Conditional::lessThan($value, $style);

$cond = Conditional::make('<=', $value, $style);
$cond = Conditional::lessThanOrEqual($value, $style);

// The "between" operator requires an array of two values.
$cond = Conditional::make('between', [10, 50], $style);
$cond = Conditional::between([10, 50], $style);

$cond = Conditional::make('!between', [10, 50], $style);
$cond = Conditional::notBetween([10, 50], $style);

// Text values
$cond = Conditional::make('=', 'Hello', $style);
$cond = Conditional::contains('Hello', $style);
$cond = Conditional::notContains('Hello', $style);
$cond = Conditional::beginsWith('Hello', $style);
$cond = Conditional::endsWith('Hello', $style);

// If you are using formulas rather than scalar values, they must begin with the "=" symbol
$cond = Conditional::make('=', '=B10+SUM(C3:D8)', $style);
```
### Expressions

If you are using formulas rather than scalar values, they must begin with the "=" symbol

```php
$cond = Conditional::expression('=SUM(C3:D8)>100', $style);
// The RC address refers to the current cell
$cond = Conditional::expression('=RC>SUM(C3:D8)', $style);
```

### Gradient fill depending on values (colorScale)

```php
// Three-color scale
// -----------------
// Conditional formatting of cells using a gradient from red through yellow to green
// Here: cells with the minimum value will be colored red,
// cells with the median value (50%) will be colored yellow,
// and cells with the maximum value will be colored green
$cond = Conditional::colorScale('ff0000', 'fff000', '00ff00');

// Two-color scale
// ---------------
// Conditional formatting of cells using a gradient from red through yellow to green
// Here: only two colors are used - from red to green
$cond = Conditional::colorScale('ff0000', '00ff00');

// Single-color scale
// ------------------
// Here: cells with the minimum value will be colored by default (white),
// and cells with the maximum value will be colored green
$cond = Conditional::colorScaleMax('00ff00');

// Here: cells with the minimum value will be colored green,
// and cells with the maximum value will be colored by default (white)
$cond = Conditional::colorScaleMin('00ff00');
```
Usually the minimum and maximum values of the range are determined automatically, 
but you can explicitly set the minimum and maximum

```php
$cond = Conditional::colorScaleNum([$min, $max], $color1, $color2);
```

### Data strip inside a cell (dataBar)

Conditional formatting of the dataBar type is used to display data bars within cells 
that visually represent values in a specific range.

```php
$cond = Conditional::dataBar($color)
    ->gradient(false) // solid bar
    ->showValue(false) // show bar without values
    ->directionRtl(true) // draw bar from right to left
;
```