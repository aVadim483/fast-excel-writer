## FastExcelWriter - Sheets

### Writing Cell Values

Usually, values is written sequentially, cell by cell, row by row. Writing to a cell moves the internal pointer
to the next cell in the row, writing a row moves the pointer to the first cell of the next row.

```php
use \avadim\FastExcelWriter\Excel;

// Create workbook
$excel = Excel::create();

// Get the sheet on which we will write data
$sheet = $excel->getSheet();

// Write data to cell by cell (the first cell is A1)
// Write number to A1 and the pointer moves to the next cell (B1)
$sheet->writeCell(123);

// Write string to B1 (pointer in C1)
$sheet->writeCell('abc');

// Pointer moves to the next cell (D1) without value writing
$sheet->nextCell();

// Now we will write value to D1 with styling
$style = [
    'format' => '#,##0.00',
    'font-color' => '#ff0000',
    'tex-align' => 'center',
];
$sheet->writeCell(0.9, $style);

// Merge cells range
$sheet->mergeCells('D1:F2');

// Write to B2 and moves pointer to C2. The pointer can only move from left to right and top to bottom
$sheet->writeTo('B2', 'value');

// Merge C3:E3, write value to merged cells and move pointer to F3  
$sheet->writeTo('C3:E3', 'other value');

```

You can write values row by row

```php
$excel = Excel::create();
$sheet = $excel->getSheet();

// Sheet::writeHeader(array header, ?array rowStyle)
// Sheet::writeRow(array row, ?array rowStyle)
// Sheet::nextRow()

// Write header values to the current row 
$sheet->writeHeader(['title1', 'title2']);

// Write header values to the current row and set format of columns A and B 
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
Also, you can define the top left cell for writing  
```php
$excel = Excel::create();
$sheet = $excel->getSheet();

// The first row is 3, all rows start in column B
$sheet->setTopLeftCell('B3');

// Write header values to the current row and set format of columns A and B 
$sheet->writeHeader(['title1' => '@integer', 'title2' => '@date'])->applyFontStyleBold();;

$data = [
    [184, '2022-01-23'],
    [835, '1971-12-08'],
    [760, '1997-05-11'],
];

foreach ($data as $rowData) {
    $sheet->writeRow($rowData);
}

```

### Direct Writing To Cells

If you need to write directly to cells, you must define the area.

```php
$excel = Excel::create();
$sheet = $excel->getSheet();

$area = $sheet->makeArea('B4:F12'); // Make write area from B4 to F12
$area = $sheet->makeArea('B4:12'); // Make write area from B4 to B12
$area = $sheet->beginArea('B4');  // Make write area from B4 to max column and max row

// Set style for single cell of area (new style will replace previous)
$area->setStyle('B4', $style1); 
// Set additional style for single cell of area (new style wil be merged with previous)
$area->addStyle('B4', $style2); 

$area->setStyle('D4:F6', $style2); // Set style for single cell of area

$area->setValue('A2:J2', 'This is demo XLSX-sheet', $headerStyle);

$area
    ->setValue('H4', 'Date', ['text-align' => 'right'])
    ->setValue('I4:J4', date('Y-m-d H:i:s'), ['font-style' => 'bold', 'format' => 'datetime', 'text-align' => 'left'])
;

// Close and write all areas
$sheet->writeAreas();

```
Other way is writing row by row or cell by cell in the defined area
```php
$excel = Excel::create();
$sheet = $excel->getSheet();

$area = $sheet->beginArea('C4');  // Make write area from C4 to max column and max row

// You can write row by row within defined area
foreach ($rows as $row) {
    $area->writeRow($row);
}
$area->nextRow();
$area->writeCell(123);
$area->writeCell(456);

// Close and write all areas
$sheet->writeAreas();

```

### Cell Formats

You can use simple and advanced formats

```php
$excel = Excel::create(['Formats']);
$sheet = $excel->getSheet();

$header = [
    'created' => '@date',
    'product_id' => '@integer',
    'quantity' => '#,##0',
    'amount' => '#,##0.00',
    'description' => '@string',
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
| @text          | @                   |
| @string        | @                   |
| @integer       | 0                   |
| @date          | YYYY-MM-DD          |
| @datetime      | YYYY-MM-DD HH:MM:SS |
| @time          | HH:MM:SS            |
| @money         | [$$]#,##0.00        |


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
You can define formula for the specified column
```php
$sheet->setColFormula('C', '=RC[-1]*0.1');

// We write values only to columns 'A' and 'B', formula to 'C' will be added automatically
$sheet->writeRow([100, 230]);
$sheet->writeRow([120, 560]);
$sheet->writeRow([130, 117]);
```

## Hyperlinks
You can insert URLs as active hyperlinks

```php
// Write URL as simple string (not hyperlink)
$sheet->writeCell('https://google.com');

// Write URL as an active hyperlink
$sheet->writeCell('https://google.com', ['hyperlink' => true]);

// Write text with an active hyperlink
$sheet->writeCell('Google', ['hyperlink' => 'https://google.com']);

```

Returns to [README.md](/README.md)
