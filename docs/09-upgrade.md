# FastExcelWriter – Upgrade Guide

This page describes important changes between major versions that may require
updating your code.

## Upgrade to version 6

The general news of v.6.0 is [Data Validation](07-validation.md) support.

### Important changes in version 6.1

* ```Sheet::setRowOptions()```, ```Sheet::setColOptions()```, ```Sheet::setRowStyles()``` and ```Sheet::setColStyles()```
are deprecated, instead of them you should use other functions: ```setRowStyle()```, ```setRowStyleArray()```,
```setRowDataStyle()```, ```setRowDataStyleArray()```, ```setColStyle()```, ```setColStyleArray()```, ```setColDataStyle()```, ```setColDataStyleArray()```
* The behavior of the ```Sheet::setRowStyle()``` and ```Sheet::setColStyle()``` has changed, they now set styles
for the entire row or column (even if they are empty)

### Important changes in version 6.9

* The namespace of the ```RichText``` class has been changed to ```avadim\FastExcelWriter\RichText```
* The namespace of the ```Style```, ```StyleManager```, and ```Font``` classes has been changed to ```avadim\FastExcelWriter\Style```
* Deprecated methods removed: ```Sheet::setColStyles()```, ```Sheet::setColOptions()```, ```Sheet::getExternalLinks()```,
```Sheet::setPageOptions()```, ```Sheet::setRowOptions()```, ```Sheet::setRowStyles()```

## Upgrade to version 5

The general news of v.5.0 is [Chart](05-charts.md) support.

### Important changes in version 5.8

Before v.5.8

```php
$sheet->writeCell(12345); // The number 12345 will be written into the cell
$sheet->writeCell('12345'); // The number 12345 will also be written here
```

In version 5.8 and later

```php
$sheet->writeCell(12345); // The number 12345 will be written into the cell
$sheet->writeCell('12345'); // Here the string '12345' will be written into the cell
```

If you want to keep the previous behavior for backward compatibility,
you should use option 'auto_convert_number' when creating a workbook.

```php
$excel = Excel::create(['Sheet1'], ['auto_convert_number' => true]);
$sheet = $excel->sheet();
$sheet->writeCell('12345'); // String '12345' will be automatically converted to a number
```

### Renamed methods in version 5.1

Some chart methods have been renamed

* ```setDataSeriesTickLabels()``` => ```setCategoryAxisLabels()```
* ```setXAxisLabel()``` => ```setCategoryAxisTitle()```
* ```getXAxisLabel()``` => ```getCategoryAxisTitle()```
* ```setYAxisLabel()``` => ```setValueAxisTitle()```
* ```getYAxisLabel()``` => ```getValueAxisTitle()```

## Upgrade to version 4

* Now the library works even faster
* Added a fluent interface for applying styles
* New methods and code refactoring

See the full list of changes in the [CHANGELOG](https://github.com/aVadim483/fast-excel-writer/blob/master/CHANGELOG.md).
