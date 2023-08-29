[![Latest Stable Version](http://poser.pugx.org/avadim/fast-excel-writer/v)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![Total Downloads](http://poser.pugx.org/avadim/fast-excel-writer/downloads)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![License](http://poser.pugx.org/avadim/fast-excel-writer/license)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![PHP Version Require](http://poser.pugx.org/avadim/fast-excel-writer/require/php)](https://packagist.org/packages/avadim/fast-excel-writer)

# FastExcelWriter

This documentation for version 4.x. Documentation for ver. 3.x see [here](/docs/ver3x.md).

Jump To:
* [Introduction](#introduction)
* [Changes in version 4](#changes-in-version-4)
* [Simple Example](#simple-example)
* [Advanced Example](#advanced-example)
* [Height And Width](#height-and-width)
* [Define Named Ranges](#define-named-ranges)
* [Adding Notes](#adding-notes)
* [Adding Images](#adding-images)
* [Workbook](/docs/01-workbook.md)
  * [Workbook settings](/docs/01-workbook.md#workbook-settings)
  * [Sets metadata of workbook](/docs/01-workbook.md#sets-metadata-of-workbook)
  * [Set Directory For Temporary Files](/docs/01-workbook.md#set-directory-for-temporary-files)
  * [Helpers methods](/docs/01-workbook.md#helpers-methods)
* [Sheets](/docs/02-sheets.md)
  * [Create, select and remove sheet](/docs/02-sheets.md#create-select-and-remove-sheet)
  * [Sheet settings](/docs/02-sheets.md#create-select-and-remove-sheet)
  * [Page settings](/docs/02-sheets.md#create-select-and-remove-sheet)
  * [Freeze Panes and Autofilter](/docs/02-sheets.md#freeze-panes-and-autofilter)
* [Writing](/docs/03-writing.md)
  * [Writing Cell Values](/docs/03-writing.md#writing-cell-values)
  * [Direct Writing To Cells](/docs/03-writing.md#direct-writing-to-cells)
  * [Cell Formats](/docs/03-writing.md#cell-formats)
  * [Formulas](/docs/03-writing.md#formulas)
  * [Hyperlinks](/docs/03-writing.md#hyperlinks)
* [Styles](/docs/04-styles.md)
  * [Cell Styles](/docs/04-styles.md#cell-styles)
  * [Row Styles](/docs/04-styles.md#row-styles)
  * [Column Styles](/docs/04-styles.md#column-styles)
  * [Other Columns Options](/docs/04-styles.md#other-columns-options)
  * [Apply Styles (The Fluent Interface)](/docs/04-styles.md#apply-styles-the-fluent-interface)
  * [Apply Borders](/docs/04-styles.md#apply-borders)
  * [Apply Fonts](/docs/04-styles.md#apply-fonts)
  * [Apply Colors](/docs/04-styles.md#apply-colors)
  * [Apply Text Styles](/docs/04-styles.md#apply-text-styles)
* [Do you want to support FastExcelWriter?](#do-you-want-to-support-fastexcelwriter)

## Introduction

This library is designed to be lightweight, super-fast and requires minimal memory usage.

This library creates Excel compatible spreadsheets in XLSX format (Office 2007+), with many features supported:
* Takes UTF-8 encoded input
* Multiple worksheets
* Supports currency/date/numeric cell formatting, formulas and active hyperlinks
* Supports basic column, row and cell styling
* You can set the height of the rows and the width of the columns (including auto width calculation)
* You can add formulas, notes and images in you XLSX-files

And if you are Laravel-developer then you can use special **Laravel** edition -
[FastExcelLaravel](https://packagist.org/packages/avadim/fast-excel-laravel).

### **FastExcelWriter** vs **PhpSpreadsheet**

**PhpSpreadsheet** is a perfect library with wonderful features for reading and writing many document formats.
**FastExcelWriter** can only write and only in XLSX format, but does it very fast 
and with minimal memory usage (if you need read XLSX-files the see **FastExcelReader** -
https://github.com/aVadim483/fast-excel-reader).

**FastExcelWriter**:
* 7-9 times faster
* uses less memory by 8-10 times
* supports writing huge 100K+ row spreadsheets

Benchmark of PhpSpreadsheet (P) and FastExcelWriter (F), spreadsheet generation without styles

| Rows x Cols | Time **P** | Time **F** | Memory **P** | Memory **F** |
|-------------|------------|------------|--------------|--------------|
| 1000 x 5    | 0.98 sec   | 0.19 sec   | 2,048 Kb     | 2,048 Kb     |
| 1000 x 25   | 4.68 sec   | 1.36 sec   | 14,336 Kb    | 2,048 Kb     |
| 5000 x 25   | 23.19 sec  | 3.61 sec   | 77,824 Kb    | 2,048 Kb     |
| 10000 x 50  | 105.8 sec  | 13.02 sec  | 256,000 Kb   | 2,048 Kb     |

## Installation

Use `composer` to install **FastExcelWriter** into your project:

```
composer require avadim/fast-excel-writer
```

Also, you can download package and include autoload file of the library:
```php
require 'path/to/fast-excel-writer/src/autoload.php';
```

## Changes In Version 4

* Now the library works even faster
* Added a fluent interface for applying styles.
* New methods and code refactoring

## Usage

You can find usage examples below or in */demo* folder

### Simple Example
```php
use \avadim\FastExcelWriter\Excel;

$data = [
    ['2003-12-31', 'James', '220'],
    ['2003-8-23', 'Mike', '153.5'],
    ['2003-06-01', 'John', '34.12'],
];

$excel = Excel::create(['Sheet1']);
$sheet = $excel->getSheet();

// Write heads
$sheet->writeRow(['Date', 'Name', 'Amount']);

// Write data
foreach($data as $rowData) {
    $rowOptions = [
        'height' => 20,
    ];
    $sheet->writeRow($rowData, $rowOptions);
}

$excel->save('simple.xlsx');
```
Also, you can download generated file to client (send to browser)
```php
$excel->download('download.xlsx');
```

### Advanced Example
```php
use \avadim\FastExcelWriter\Excel;

$head = ['Date', 'Name', 'Amount'];
$data = [
    ['2003-12-31', 'James', '220'],
    ['2003-8-23', 'Mike', '153.5'],
    ['2003-06-01', 'John', '34.12'],
];
$headStyle = [
    'font' => [
        'style' => 'bold'
    ],
    'text-align' => 'center',
    'vertical-align' => 'center',
    'border' => 'thin',
    'height' => 24,
];

$excel = Excel::create(['Sheet1']);
$sheet = $excel->getSheet();

// Write the head row (sets style via array)
$sheet->writeHeader($head, $headStyle);

// The same result with new fluent interface
$sheet->writeHeader($head)
    ->applyFontStyleBold()
    ->applyTextAlign('center', 'center')
    ->applyBorder(Style::BORDER_STYLE_THIN)
    ->applyRowHeight(24);

// Sets columns options - format and width (the first way)
$sheet
    ->setColFormats(['@date', '@text', '0.00'])
    ->setColWidths([12, 14, 5]);

// The seconds way to set columns options
$sheet
    // column and options
    ->setColOptions('A', ['format' => '@date', 'width' => 12])
    // column letter in lower case
    ->setColOptions('b', ['format' => '@text', 'width' => 24])
    // column can be specified by number
    ->setColOptions(3, ['format' => '0.00', 'width' => 15, 'color' => '#090'])
;

// The third way - all options in multilevel array (first level keys point to columns)
$sheet
    ->setColOptions([
        'A' => ['format' => '@date', 'width' => 12],
        'B' => ['format' => '@text', 'width' => 24],
        'C' => ['format' => '0.00', 'width' => 15, 'color' => '#090'],
    ]);

$rowNum = 1;
foreach($data as $rowData) {
    $rowOptions = [
        'height' => 20,
    ];
    if ($rowNum % 2) {
        $rowOptions['fill-color'] = '#eee';
    }
    $sheet->writeRow($rowData, $rowOptions);
}

$excel->save('simple.xlsx');
```


### Height And Width

```php
// Set height of row 2 to 33
$sheet->setRowHeight(2, 33);
// Set heights of several rows
$sheet->setRowHeights([1 => 20, 2 => 33, 3 => 40]);
// Write row data and set height
$sheet->writeRow($rowData, ['height' => 20]);

// Set width of column D to 24
$this->setColWidth('D', 24);
$this->setColOptions('D', ['width' => 24]);
// Set auto width
$this->setColWidth('D', 'auto');
$this->setColWidthAuto('D');
$this->setColOptions('D', ['width' => 'auto']);

// Set width of specific columns
$sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);
$colOptions = [
    'B' => ['width' => 10], 
    'C' => ['width' => 'auto'], 
    'E' => ['width' => 30], 
    'F' => ['width' => 40],
];
$sheet->setColOptions($colOptions);
// Set width of columns from 'A'
$sheet->setColWidths([10, 20, 30, 40]);

```

### Define Named Ranges

FastExcelWriter supports _named ranges_ and does not support _named formulae_.
A _named ranges_ provides a name reference to a cell or a range of cells.
All _named ranges_ are added to the workbook so all names must be unique, but you can define _named ranges_ in a sheet or in a workbook.

Also range names must start with a letter or underscore, have no spaces, and be no longer than 255 characters.

```php
$excel = Excel::create();
$excel->setFileName($outFileName);
$sheet = $excel->getSheet();

// Named a single cell
$sheet->addNamedRange('B2', 'cell_name');

// Named range in a sheet
$sheet->addNamedRange('c2:e3', 'range_name');

// Add named range in a workbook (sheet name required)
$excel->addNamedRange('Sheet1!A1:F5', 'A1_F5');

```

###  Adding Notes

There are currently two types of comments in Excel - **comments** and **notes** 
(see [The difference between threaded comments and notes](https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3)).
Notes are old style comments in Excel (text on a light yellow background). 
You can add notes to any cells using method ```addNote()```

```php

$sheet->writeCell('Text to A1');
$sheet->addNote('A1', 'This is a note for cell A1');

$sheet->writeCell('Text to B1')->addNote('This is a note for B1');
$sheet->writeTo('C4', 'Text to C4')->addNote('Note for C1');

// If you specify a range of cells, then the note will be added to the left top cell
$sheet->addNote('E4:F8', "This note\nwill added to E4");

// You can split text into multiple lines
$sheet->addNote('D7', "Line 1\nLine 2");

```

You can change some note options. Allowed options of a note are:
* **width** - default value is ```'96pt'```
* **height** - default value is ```'55.5pt'```
* **fill_color** - default value is ```'#FFFFE1'```
* **show** - default value is ```false```

```php

$sheet->addNote('A1', 'This is a note for cell A1', ['width' => '200pt', 'height' => '100pt', 'fill_color' => '#ffcccc']);

// Parameters "width" and "height" can be numeric, by default these values are in points
// The "fill_color" parameter can be shortened
$noteStyle = [
    'width' => 200, // equivalent to '200pt'
    'height' => 100, // equivalent to '100pt'
    'fill_color' => 'fcc', // equivalent to '#ffcccc'
];
$sheet->writeCell('Text to B1')->addNote('This is a note for B1', $noteStyle);

// This note is visible when the Excel workbook is displayed
$sheet->addNote('C8', 'This note is always visible', ['show' => true]);
```

###  Adding Images

```php
// Insert an image to the cell A1
$sheet->addImage('A1', 'path/to/file');

// Insert an image to the cell B2 and set with to 150 pixels (height will change proportionally)
$sheet->addImage('B2', 'path/to/file', ['width' => 150]);

// Set height to 150 pixels (with will change proportionally)
$sheet->addImage('C3', 'path/to/file', ['height' => 150]);

// Set size in pixels
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150]);

```

## Do you want to support FastExcelWriter?

if you find this package useful you can support and donate to me for a cup of coffee:

* USDT (TRC20) TSsUFvJehQBJCKeYgNNR1cpswY6JZnbZK7
* USDT (ERC20) 0x5244519D65035aF868a010C2f68a086F473FC82b
* ETH 0x5244519D65035aF868a010C2f68a086F473FC82b

Or just give me a star on GitHub :)