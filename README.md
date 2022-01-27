# FastExcelWriter

This library is designed to be lightweight, super fast and have minimal memory usage. 
Partially based on https://github.com/mk-j/PHP_XLSXWriter, but advanced and improved.

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

Also you can download package and include autoload file of the library:
```php
require 'path/to/fast-excel-writer/src/autoload.php';
```

## Usage

You can find usage examples below or in */demo* folder

### Simple example
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
### Formulas

Formulas must start with '='. If you want to write the formula as text, use a backslash. 
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

| simple formats | format code |
| ---------- | ---- |
| text     | @ |
| string   | @ |
| integer  | 0 |
| date     | YYYY-MM-DD |
| datetime | YYYY-MM-DD HH:MM:SS |
| time     | HH:MM:SS |
| money    | [$$]#,##0.00 |

### Basic cell styles

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

| key          | allowed values |
| ------------ | ---- |
| name         | Arial, Times New Roman, Courier New, Comic Sans MS |
| size         | 8, 9, 10, 11, 12 ... |
| style        | bold, italic, underline, strikethrough or multiple ie: 'bold,italic' |

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
|----------------|------------------------------------------------ |
| color          | #RRGGBB, ie: '#ff99cc' or '#f9c'                |
| fill           | #RRGGBB, ie: '#eeffee' or '#efe'                |
| text-align     | 'general', 'left', 'right', 'justify', 'center' |
| vertical-align | 'bottom', 'center', 'distributed'               |
| text-wrap      | true, false                                     |

## Other options
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