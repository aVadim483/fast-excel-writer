[![GitHub Release](https://img.shields.io/github/v/release/aVadim483/fast-excel-writer)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![Packagist Downloads](https://img.shields.io/packagist/dt/avadim/fast-excel-writer?color=%2300aa00)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![License](http://poser.pugx.org/avadim/fast-excel-writer/license)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![PHP Version Require](http://poser.pugx.org/avadim/fast-excel-writer/require/php)](https://packagist.org/packages/avadim/fast-excel-writer)

<table border="0">
<tr>
<td valign="top"><img height="240px" src="logo/logo2-680.jpg" alt="FastExcelWriter Logo"></td>
<td valign="top">
<p align="center">
<img height="120px" src="logo/01-feature-fast.jpg" alt="fastest">
<img height="120px" src="logo/02-feature-memory.jpg" alt="memory saving">
<img height="120px" src="logo/03-feature-worksheets.jpg" alt="multiple worksheet">
<br>
<img height="120px" src="logo/04-feature-protection.jpg" alt="book and sheet protection">
<img height="120px" src="logo/05-feature-charts.jpg" alt="multiple charts">
<img height="120px" src="logo/06-feature-styling.jpg" alt="styling and image insertion">
</p>

<h1 align="center"><b>FastExcelWriter</b></h1>
</td>
</tr>
</table>

**FastExcelWriter** is a part of the **FastExcelPhp Project** which consists of

* [FastExcelWriter](https://packagist.org/packages/avadim/fast-excel-writer) - to create Excel spreadsheets
* [FastExcelReader](https://packagist.org/packages/avadim/fast-excel-reader) - to read Excel spreadsheets
* [FastExcelTemplator](https://packagist.org/packages/avadim/fast-excel-templator) - to generate Excel spreadsheets from XLSX templates
* [FastExcelLaravel](https://packagist.org/packages/avadim/fast-excel-laravel) - special **Laravel** edition

## Introduction

This library is designed to be lightweight, super-fast and requires minimal memory usage.

**FastExcelWriter** creates Excel compatible spreadsheets in XLSX format (Office 2007+), with many features supported:

* Takes UTF-8 encoded input
* Multiple worksheets
* Supports currency/date/numeric cell formatting, formulas and active hyperlinks
* Supports most styling options for cells, rows, columns - colors, borders, fonts, etc.
* You can set the height of the rows and the width of the columns (including auto width calculation)
* You can add formulas, notes and images in you XLSX-files
* Supports workbook and sheet protection with/without passwords 
* Supports page settings - page margins, page size
* Inserting multiple charts

Jump To:
* [Changes in version 4](#changes-in-version-4)
* [Changes in version 5](#changes-in-version-5)
* [Simple Example](#simple-example)
* [Advanced Example](#advanced-example)
* [Adding Notes](#adding-notes)
* [Adding Images](#adding-images)
* [Workbook](/docs/01-workbook.md)
  * [Workbook settings](/docs/01-workbook.md#workbook-settings)
  * [Sets metadata of workbook](/docs/01-workbook.md#sets-metadata-of-workbook)
  * [Set Directory For Temporary Files](/docs/01-workbook.md#set-directory-for-temporary-files)
  * [Helpers methods](/docs/01-workbook.md#helpers-methods)
* [Sheets](/docs/02-sheets.md)
  * [Create, select and remove sheet](/docs/02-sheets.md#create-select-and-remove-sheet)
  * [Sheet settings](/docs/02-sheets.md#sheet-settings)
  * [Page settings](/docs/02-sheets.md#page-settings)
  * [Row's settings](/docs/02-sheets.md#rows-settings)
  * [Column's settings](/docs/02-sheets.md#columns-settings)
  * [Group/outline rows and columns](/docs/02-sheets.md#groupoutline-rows-and-columns)
  * [Define Named Ranges](/docs/02-sheets.md#define-named-ranges)
  * [Freeze Panes and Autofilter](/docs/02-sheets.md#freeze-panes-and-autofilter)
  * [Setting Active Cells](/docs/02-sheets.md#setting-active-cells)
  * [Print Settings](/docs/02-sheets.md#print-settings)
* [Writing](/docs/03-writing.md)
  * [Writing Row by Row vs Direct](/docs/03-writing.md#writing-row-by-row-vs-direct)
  * [Direct Writing To Cells](/docs/03-writing.md#direct-writing-to-cells)
  * [Writing Cell Values](/docs/03-writing.md#writing-cell-values)
  * [Merging Cells](/docs/03-writing.md#merging-cells)
  * [Cell Formats](/docs/03-writing.md#cell-formats)
  * [Formulas](/docs/03-writing.md#formulas)
  * [Hyperlinks](/docs/03-writing.md#hyperlinks)
  * [Using Rich Text](/docs/03-writing.md#using-rich-text)
* [Styles](/docs/04-styles.md)
  * [Cell Styles](/docs/04-styles.md#cell-styles)
  * [Row Styles](/docs/04-styles.md#row-styles)
  * [Column Styles](/docs/04-styles.md#column-styles)
  * [Other Columns Options](/docs/04-styles.md#other-columns-options)
  * [Apply Styles (The Fluent Interface)](/docs/04-styles.md#apply-styles--the-fluent-interface-)
  * [Apply Borders](/docs/04-styles.md#apply-borders)
  * [Apply Fonts](/docs/04-styles.md#apply-fonts)
  * [Apply Colors](/docs/04-styles.md#apply-colors)
  * [Apply Text Styles](/docs/04-styles.md#apply-text-styles)
* [Charts](/docs/05-charts.md)
  * [Simple usage](/docs/05-charts.md#simple-usage-of-chart)
  * [Combo charts](/docs/05-charts.md#combo-charts)
  * [Multiple charts](/docs/05-charts.md#multiple-charts)
  * [Chart types](/docs/05-charts.md#chart-types)
  * [Useful Chart Methods](/docs/05-charts.md#useful-chart-methods)
* [Protection of workbook and sheets](/docs/06-protection.md)
  * [Workbook protection](/docs/06-protection.md#workbook-protection)
  * [Sheet protection](/docs/06-protection.md#sheet-protection)
  * [Cells locking/unlocking](/docs/06-protection.md#cells-lockingunlocking)
* [FastExcelWriter vs PhpSpreadsheet](#fastexcelwriter-vs-phpspreadsheet)
* [Do you want to support FastExcelWriter?](#do-you-want-to-support-fastexcelwriter)


## Installation

Use `composer` to install **FastExcelWriter** into your project:

```
composer require avadim/fast-excel-writer
```

## Changes In Version 4

* Now the library works even faster
* Added a fluent interface for applying styles.
* New methods and code refactoring

## Changes In Version 5

* The general news is Chart support

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
$sheet = $excel->sheet();

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
$sheet = $excel->sheet();

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



###  Adding Notes

There are currently two types of comments in Excel - **comments** and **notes** 
(see [The difference between threaded comments and notes](https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3)).
Notes are old style comments in Excel (text on a light yellow background). 
You can add notes to any cells using method ```addNote()```

```php

$sheet1->writeCell('Text to A1');
$sheet1->addNote('A1', 'This is a note for cell A1');

$sheet1->writeCell('Text to B1')->addNote('This is a note for B1');
$sheet1->writeTo('C4', 'Text to C4')->addNote('Note for C1');

// If you specify a range of cells, then the note will be added to the left top cell
$sheet1->addNote('E4:F8', "This note\nwill added to E4");

// You can split text into multiple lines
$sheet1->addNote('D7', "Line 1\nLine 2");

```

You can change some note options. Allowed options of a note are:
* **width** - default value is ```'96pt'```
* **height** - default value is ```'55.5pt'```
* **fill_color** - default value is ```'#FFFFE1'```
* **show** - default value is ```false```

```php

$sheet1->addNote('A1', 'This is a note for cell A1', ['width' => '200pt', 'height' => '100pt', 'fill_color' => '#ffcccc']);

// Parameters "width" and "height" can be numeric, by default these values are in points
// The "fill_color" parameter can be shortened
$noteStyle = [
    'width' => 200, // equivalent to '200pt'
    'height' => 100, // equivalent to '100pt'
    'fill_color' => 'fcc', // equivalent to '#ffcccc'
];
$sheet1->writeCell('Text to B1')->addNote('This is a note for B1', $noteStyle);

// This note is visible when the Excel workbook is displayed
$sheet1->addNote('C8', 'This note is always visible', ['show' => true]);
```

Also, you can use rich text in notes

```php
$richText = new \avadim\FastExcelWriter\RichText('here is <c=f00>red</c> and <c=00f>blue</c> text');
$sheet1->addNote('C8', $richText);
```

For more information on using rich text, see here: [Using Rich Text](/docs/03-writing.md#using-rich-text)

###  Adding Images

```php
// Insert an image to the cell A1
$sheet1->addImage('A1', 'path/to/file');

// Insert an image to the cell B2 and set with to 150 pixels (height will change proportionally)
$sheet1->addImage('B2', 'path/to/file', ['width' => 150]);

// Set height to 150 pixels (with will change proportionally)
$sheet1->addImage('C3', 'path/to/file', ['height' => 150]);

// Set size in pixels
$sheet1->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150]);

```

## **FastExcelWriter** vs **PhpSpreadsheet**

**PhpSpreadsheet** is a perfect library with wonderful features for reading and writing many document formats.
**FastExcelWriter** can only write and only in XLSX format, but does it very fast
and with minimal memory usage.

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

## Do you want to support FastExcelWriter?

if you find this package useful you can support and donate to me for a cup of coffee:

* USDT (TRC20) TSsUFvJehQBJCKeYgNNR1cpswY6JZnbZK7
* USDT (ERC20) 0x5244519D65035aF868a010C2f68a086F473FC82b
* ETH 0x5244519D65035aF868a010C2f68a086F473FC82b

Or just give me a star on GitHub :)