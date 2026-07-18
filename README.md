[![GitHub Release](https://img.shields.io/github/v/release/aVadim483/fast-excel-writer)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![Packagist Downloads](https://img.shields.io/packagist/dt/avadim/fast-excel-writer?color=%23aa00aa)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![GitHub License](https://img.shields.io/github/license/aVadim483/fast-excel-writer)](https://packagist.org/packages/avadim/fast-excel-writer) 
[![Static Badge](https://img.shields.io/badge/php-%3E%3D7.4-005fc7)](https://packagist.org/packages/avadim/fast-excel-writer)

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

<h1 align="center"><b>FastExcelWriter v.6</b></h1>
</td>
</tr>
</table>

**FastExcelWriter** is a part of the **FastExcelPhp Project** which consists of

* [FastExcelWriter](https://packagist.org/packages/avadim/fast-excel-writer) - to create Excel spreadsheets
* [FastExcelReader](https://packagist.org/packages/avadim/fast-excel-reader) - to read Excel spreadsheets
* [FastExcelTemplator](https://packagist.org/packages/avadim/fast-excel-templator) - to generate Excel spreadsheets from XLSX templates
* [FastExcelLaravel](https://packagist.org/packages/avadim/fast-excel-laravel) - special **Laravel** edition

## Introduction

Lightweight, powerful and very fast XLSX Spreadsheet Writer in pure PHP. This library is designed to be super-fast and requires minimal memory usage.

**FastExcelWriter** creates spreadsheets in XLSX format, compatible with MS Excel (Office 2007+), LibreOffice, OpenOffice and others, 
with many features supported:

* Takes UTF-8 encoded input
* Multiple worksheets
* Supports currency/date/numeric cell formatting, formulas and active hyperlinks
* Supports most styling options for cells, rows, columns - colors, borders, fonts, etc.
* You can set the height of the rows and the width of the columns (including auto width calculation)
* You can add formulas, notes and images in your XLSX-files
* Supports workbook and sheet protection with/without passwords 
* Supports page settings - page margins, page size
* Inserting multiple charts
* Supports data validations and conditional formatting

Currently, the required PHP minimum version is PHP 7.4

## Installation

Use `composer` to install **FastExcelWriter** into your project:

```
composer require avadim/fast-excel-writer
```

## Quick Start

```php
use \avadim\FastExcelWriter\Excel;
use \avadim\FastExcelWriter\Style\Style;

$data = [
    ['2003-12-31', 'James', '220'],
    ['2003-8-23', 'Mike', '153.5'],
    ['2003-06-01', 'John', '34.12'],
];

$excel = Excel::create(['Sheet1']);
$sheet = $excel->sheet();

// Write the head row and set styles via fluent interface
$sheet->writeHeader(['Date', 'Name', 'Amount'])
    ->applyFontStyleBold()
    ->applyBorder(Style::BORDER_STYLE_THIN);

// Set column formats and widths
$sheet
    ->setColFormats(['@date', '@text', '0.00'])
    ->setColWidths([12, 14, 8]);

// Write data
foreach ($data as $rowData) {
    $sheet->writeRow($rowData);
}

// Save to a file
$excel->save('simple.xlsx');

// ...or download generated file to client (send to browser)
// $excel->download('simple.xlsx');
```

## Documentation

Read more in the documentation [here](/docs/index.md) or [here](https://fast-excel-writer.readthedocs.io/en/latest/). 
Or check out the [API reference](/docs/90-api-reference.md). 
Also, you can see examples in ```/demo``` folder.

Upgrading from an older version? See the [upgrade guide](/docs/09-upgrade.md).

Changelog is [here](CHANGELOG.md).

## **FastExcelWriter** vs **PhpSpreadsheet**

**PhpSpreadsheet** is a perfect library with wonderful features for reading and writing many document formats.
**FastExcelWriter** can only write and only in XLSX format, but does it very fast
and with minimal memory usage.

**FastExcelWriter**:
* 7-9 times faster
* uses less memory by 8-10 times
* supports writing huge 100K+ row spreadsheets

Benchmark of PhpSpreadsheet (generation without styles)

| Rows x Cols | Time      | Memory     |
|-------------|-----------|------------|
| 1000 x 5    | 0.98 sec  | 2,048 Kb   |
| 1000 x 25   | 4.68 sec  | 14,336 Kb  |
| 5000 x 25   | 23.19 sec | 77,824 Kb  |
| 10000 x 50  | 105.8 sec | 256,000 Kb |

Benchmark of FastExcelWriter (generation without styles)

| Rows x Cols | Time      | Memory   |
|-------------|-----------|----------|
| 1000 x 5    | 0.19 sec  | 2,048 Kb |
| 1000 x 25   | 1.36 sec  | 2,048 Kb |
| 5000 x 25   | 3.61 sec  | 2,048 Kb |
| 10000 x 50  | 13.02 sec | 2,048 Kb |

## Do you want to support FastExcelWriter?

if you find this package useful you can support and donate to me for a cup of coffee:

* USDT (TRC20) TSsUFvJehQBJCKeYgNNR1cpswY6JZnbZK7
* USDT (ERC20) 0x5244519D65035aF868a010C2f68a086F473FC82b
* ETH 0x5244519D65035aF868a010C2f68a086F473FC82b

Or just give me a star on GitHub :)
