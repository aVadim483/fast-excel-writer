# FastExcelWriter

Jump To:
* [Introduction](#introduction)
* [Installation](#installation)
* [Simple Example](#simple-example)
* [Writing Cell Values](#writing-cell-values)
* [Cell Formats](#cell-formats)
* [Basic Cell Styles](#basic-cell-styles)
* [Formulas](#formulas)
* [Set Directory For Temporary Files](#set-directory-for-temporary-files)

## Introduction

This library is designed to be lightweight, superfast and have minimal memory usage. 

This library creates Excel compatible spreadsheets in XLSX format (Office 2007+), with just basic features supported:
* takes UTF-8 encoded input
* multiple worksheets
* supports currency/date/numeric cell formatting, simple formulas
* supports basic column, row and cell styling


**FastExcelWriter** vs **PhpSpreadsheet**

**PhpSpreadsheet** is a perfect library with wonderful features for reading and writing many document formats.
**FastExcelWriter** can only write and only in xlsx format, but does it very fast 
and with minimal memory usage.

**FastExcelWriter**:
* 7-9 times faster
* uses less memory by 8-10 times
* supports writing huge 100K+ row spreadsheets

By the way, **FastExcelReader** also exists - https://github.com/aVadim483/fast-excel-reader

## Installation

Use `composer` to install **FastExcelWriter** into your project:

```
composer require avadim/fast-excel-writer
```

Also, you can download package and include autoload file of the library:
```php
require 'path/to/fast-excel-writer/src/autoload.php';
```

## Usage

You can find usage examples below or in */demo* folder

### Simple Example
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

$sheet->writeRow($head, $headStyle);

$sheet
    ->setColFormats(['date', 'string', '0.00'])
    ->setColWidths([12, 14, 5]);

$rowNum = 1;
foreach($data as $rowData) {
    $rowOptions = [
        'height' => 20,
    ];
    if ($rowNum % 2) {
        $rowOptions['fill'] = '#eee';
    }
    $sheet->writeRow($rowData);
}

$excel->save('simple.xlsx');
```
Also, you can download generated file to client (send to browser)
```php
$excel->output('download.xlsx');
```

### Writing Cell Values

Usually, values is written sequentially, cell by cell, row by row. Writing to a cell moves the internal pointer 
to the next cell in the row, writing a row moves the pointer to the first cell of the next row.

```php
use \avadim\FastExcelWriter\Excel;

$excel = Excel::create();
$sheet = $excel->getSheet();

// Sheet::writeCell(mixed value, ?array styles)
// Sheet::writeTo(string address, mixed value, ?array styles)
// Sheet::nextCell()

// Write number to A1 and pointer moves to the next cell (B1)
$sheet->writeCell(123);
// Write string to B1 (pointer in C1)
$sheet->writeCell('abc');
// Pointer moves to the next cell (D1) without value writing
$sheet->nextCell();
$style = [
    'color' => '#ff0000',
    'format' => '#,##0.00',
    'align' => 'center',
];
// Write to D1 value and style
$sheet->writeCell(0.9, $style);
// Merge cells range
$sheet->mergeCells('D1:F2');
// Write to B2 and moves pointer to C2. The pointer can only move from left to right and top to bottom
$sheet->writeTo('B2', 'value');
// Merge C3:E3, write value to merged cells and move pointer to F3  
$sheet->writeTo('C3:E3', 'other value');
```

You can write values to rows

```php
// Sheet::writeHeader(array header, ?array rowStyle)
// Sheet::writeRow(array row, ?array rowStyle)
// Sheet::nextRow()

// Write values to the current row and set format of columns A and B 
$sheet->writeHeader(['title1' => '@integer', 'title2' => '@date']);

$data = [
    [184, '2022-01-23'],
    [835, '1971-12-08'],
    [760, '1997-05-11'],
];

foreach ($data as $rowData) {
    $sheet->writeRow($rowData);
}


```

### Cell Formats

You can use simple and advanced formats

```php
$excel = new \avadim\FastExcelWriter\Excel(['Formats']);
$sheet = $excel->getSheet();

$header = [
    'created' => 'date',
    'product_id' => 'integer',
    'quantity' => '#,##0',
    'amount' => '#,##0.00',
    'description' => 'string',
    'tax' => '[$$]#,##0.00;[RED]-[$$]#,##0.00',
];
$data = [
    ['2015-01-01', 873, 1, 44.00, 'misc', '=D2*0.05'],
    ['2015-01-12', 324, 2, 88.00, 'none', '=D3*0.15'],
];

$sheet->writeHeader($header);
foreach($data as $row) {
    $sheet->writeRow($row );
}

$excel->save('formats.xlsx');
```

Simple cell formats map to more advanced cell formats

| simple formats | format code         |
|----------------|---------------------|
| text           | @                   |
| string         | @                   |
| integer        | 0                   |
| date           | YYYY-MM-DD          |
| datetime       | YYYY-MM-DD HH:MM:SS |
| time           | HH:MM:SS            |
| money          | [$$]#,##0.00        |

### Basic Cell Styles

Font settings
```php
use \avadim\FastExcelWriter\Style;

$style = [
    Style::FONT => [
        Style::FONT_NAME => 'Arial',
        Style::FONT_SIZE => 14,
        Style::FONT_STYLE => Style::FONT_STYLE_BOLD,
    ]
];
```

| key           | allowed values                                                       |
|---------------|----------------------------------------------------------------------|
| name          | Arial, Times New Roman, Courier New, Comic Sans MS                   |
| size          | 8, 9, 10, 11, 12 ...                                                 |
| style         | bold, italic, underline, strikethrough or multiple ie: 'bold,italic' |

Border settings
```php
use \avadim\FastExcelWriter\Style;

// simple border style
$style1 = [
    Style::BORDER => Style::BORDER_THIN
];

// border style with color
$style2 = [
    Style::BORDER => [
        Style::BORDER_ALL => [
            Style::BORDER_STYLE => Style::BORDER_THICK,
            Style::BORDER_COLOR => '#f00',
        ]
    ]
];

// extra border style
$style3 = [
    Style::BORDER => [
        Style::BORDER_TOP => Style::BORDER_NONE,
        Style::BORDER_LEFT => [
            Style::BORDER_STYLE => Style::BORDER_THICK,
            Style::BORDER_COLOR => '#f9009f',
        ],
        Style::BORDER_RIGHT => [
            Style::BORDER_STYLE => Style::BORDER_MEDIUM_DASH_DOT,
            Style::BORDER_COLOR => '#f00',
        ],
        Style::BORDER_BOTTOM => [
            Style::BORDER_STYLE => Style::BORDER_DOUBLE,
        ],
    ]
];
```
Other style settings

| style          | allowed values                                  |
|----------------|-------------------------------------------------|
| color          | #RRGGBB, ie: '#ff99cc' or '#f9c'                |
| fill           | #RRGGBB, ie: '#eeffee' or '#efe'                |
| text-align     | 'general', 'left', 'right', 'justify', 'center' |
| vertical-align | 'bottom', 'center', 'distributed'               |
| text-wrap      | true, false                                     |


### Formulas

Formulas must start with '='. If you want to write the formula as a text, use a backslash.
Setting the locale allows the use of national language function names.
You can use both A1 and R1C1 notations in formulas

```php
use \avadim\FastExcelWriter\Excel;

$excel = Excel::create(['Formulas']);
$sheet = $excel->getSheet();

// Set Russian locale
$excel->setLocale('ru');

$headRow = [];

$sheet->writeRow([1, random_int(100, 999), '=RC[-1]*0.1']);
$sheet->writeRow([2, random_int(100, 999), '=RC[-1]*0.1']);
$sheet->writeRow([3, random_int(100, 999), '=RC[-1]*0.1']);

$totalRow = [
    'Total',
    '=SUM(B1:B3)', // English function name
    '=СУММ(C1:C3)', // You can use Russian function name because the locale is 'ru'
];

$sheet->writeRow($totalRow);

$excel->save('formulas.xlsx');

```


## Set Directory For Temporary Files
The library uses temporary files to generate the XLSX-file. If not specified, they are created in the system temporary directory
or in the current execution directory. But you can set the directory for temporary files.

```php
use \avadim\FastExcelWriter\Excel;

Excel::setTempDir('/path/to/temp/dir'); // use this call before Excel::create()
$excel = Excel::create();

// Or alternative variant

$excel = Excel::create('SheetName', ['temp_dir' => '/path/to/temp/dir']);
```

## Want to support FastExcelWriter?

if you find this package useful you can support and donate to me https://www.paypal.me/VShemarov
Or just give me star on Github :)