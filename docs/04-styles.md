## FastExcelWriter - Styles

You can set styles for the entire workbook, for sheets, for individual cells, as well as for rows and columns.
The style of each cell can be determined by the combination of all these styles.

### Cell Styles

```php
$style = [
    'format' => '#,##0.00',
    'font-color' => '#ff0000',
    'text-align' => 'center',
];
$sheet1->writeCell(0.9, $style);
$sheet1->writeTo('B4', $value, $style);

// Sets style to specified cell
$sheet1->setStyle('C8', $style);

// Begin an area for direct write
$area = $sheet1->beginArea();
$area
    ->setValue('C10', 1234, $otherStyle)
    ->setValue('E12:K12', 'This is string', $alterStyle);

```

### Row Styles

```php
$rowStyle = [
    Style::FILL_COLOR => '#ff99ff',
    Style::BORDER => [
        Style::BORDER_TOP => [
            Style::BORDER_STYLE => Style::BORDER_THICK,
            Style::BORDER_COLOR => '#f00',
        ]
    ]
];
// Write row data and set row styles
$sheet1->writeRow($rowData, $rowStyle);

// Set options for several rows 
$sheet1->setRowStyles('3', $style);

$rowStyles = [
    3 => ['fill-color' = '#cff'], // options for row 3 
    4 => ['fill-color' = '#ccc', 'height' = 20], // options for row 4
];

// Set styles to the specified row 
$sheet1->setRowStyles($rowStyles);

// Set options for range of rows 
$sheet1->setRowStyles('2:5', ['fill-color' = '#f00']);

```

You can set the style for the entire sheet row or only for those cells in the row where data is written.
```php
// Style are applied to the entire sheet row
$sheet->setRowStyle(3, ['height' = 20]);
$sheet->setRowStyle('2:5', ['font-color' = '#f00']);
$sheet->setRowStyleArray([3 => $style1, 5 => $style2]);

// Set the style only for non-empty cells in a row
$sheet->setRowDataStyle(3, ['height' = 20]);
$sheet->setRowDataStyle('2:5', ['font-color' = '#f00']);
$sheet->setRowDataStyleArray([3 => $style1, 5 => $style2]);
```


### Row And Cell Styles

The first argument of the function ```writeRow()``` is an array of values, the 2nd is the style of the row, and the 3rd is the styles of the cells in the row.

```php
$rowValues = [
    'text',
    'http://google.com',
    123456,
];
$rowStyle = [];
$cellStyles = [
    [], // style for the first cell,
    ['hyperlink' => true], // 2nd cell
    [], // 3rd cell
];

$sheet1->writeRow($rowValues, $rowStyle, $cellStyles);

```

Also, you can set styles for the specified cells

```php
$cellStyles = [
    'B' => ['hyperlink' => true],
];
$sheet1->writeRow($rowValues, [], $cellStyles);
```

### Column Styles

You can define column style and format with method ```writeHeader()```.
The code below will do it:
* Writes to cells of the current row values 'column title 1', 'column title 2', 'column title 3'
* Sets for this row bold font and thin border style (the default color of  borders is #000000)
* Set styles, widths and formats for the columns 'A', 'B' and 'C'

```php
$headValues = [
    // <cell value> => [<column style>]
    ['column title 1' => ['format' => '@text', 'width' => 20, 'fill-color' => '#ccf']],
    ['column title 2' => ['format' => '@integer', 'width' => 12, 'font-color' => '#009']],
    ['column title 3' => ['text-wrap' => true, 'fill-color' => '#ccf']],
];
$headStyles = [
    'font-style' => 'bold',
    'border-style' => 'thin',
];
$sheet1->writeHeader($headValues, $headStyles);

```
You can set styles for specified columns

```php
$sheet1->setColStyle('C', $colStyle);
$sheet1->setColWidth('E', 32);
$sheet1->setColFormat('K', '@date');

```

You can set the style for the entire sheet column or only for those cells in the column where data is written.
```php
// Style are applied to the entire sheet column
$sheet->setColStyle('B', $style); // style for cells of column 'B'
$sheet->setColStyle(2, $style); // column 'B' has a number 2
$sheet->setColStyle('C:F', $style); // style for range of columns
$sheet->setColStyle(['A', 'B', 'C'], $style); // options for several columns 'A', 'B' and 'C'
$sheet->setColStyleArray(['B' => ['width' = 20], 'C' => ['font-color' = '#f00']]);

// Set the style only for non-empty cells in a column
$sheet->setColDataStyle('B', ['width' = 20]);
$sheet->setColDataStyle(2, ['width' = 20]);
$sheet->setColDataStyle('B:D', ['width' = 'auto']);
$sheet->setColDataStyle(['A', 'B', 'C'], $style);
$sheet->setColDataStyleArray(['B' => $style1, 'C' => $style2]);
```

### Other Columns Options

```php
// Set style and width for specified column
$sheet1->setColDataStyle('K', ['text-wrap' => true, 'width' => 32]);

// Set widths of columns from the first (A)
$sheet1->setColWidths([5, 16, 16, 'auto']);

// Set width of the column
$sheet1->setColWidth(['G', 'H', 'J'], 14);

// Set formats of columns from the first (A); null - default format
$sheet1->setColFormats([null, '@', '@', '@date', '0', '0.00', '@money', '@money']);

```

### Apply Styles (The Fluent Interface)

Methods that start with ```'apply...'``` are applied to either the cell or group of cells where the data was written.

```php
// Create Excel workbook
$excel = Excel::create();

// Get the first sheet;
$sheet = $excel->getSheet();

// The background color will be applied to all changed cells in the row
$sheet->writeRow(['foo', 'bar'])->applyBgColor('#f99');

// The background color will only be applied to the last changed cell
$sheet->writeCell('abc')->applyBgColor('#9f9');

// The background color will only be applied to cell C3
$sheet->writeTo('C3', 'edf')->applyBgColor('#cc99ff');

// Select the specified range and apply outer and inner cell borders for it
$sheet->withRange('B4:D5')->applyBgColor('#cff')->applyBorderOuter(Style::BORDER_DOUBLE)->applyBorderInner(Style::BORDER_DOTTED);

```

#### Apply Borders
* applyBorder(string $style, ?string $color = '#000000')
* applyBorderLeft(string $style, ?string $color = '#000000')
* applyBorderRight(string $style, ?string $color = '#000000')
* applyBorderTop(string $style, ?string $color = '#000000')
* applyBorderBottom(string $style, ?string $color = '#000000')
* applyBorderOuter(string $style, ?string $color = '#000000')
* applyBorderInner(string $style, ?string $color = '#000000')

#### Apply Fonts
* applyFont(string $fontName, ?int $fontSize = null, ?string $fontStyle = null, ?string $fontColor = null)
* applyFontName(string $fontName)
* applyFontSize(float $fontSize)
* applyFontStyle(string $fontStyle)
* applyFontColor(string $fontColor)
* applyFontStyleBold()
* applyFontStyleItalic()
* applyFontStyleStrikethrough()
* applyFontStyleUnderline(?bool $double = false)

#### Apply Colors
* applyColor(string $color)
* applyTextColor(string $color)
* applyFillColor(string $color)
* applyBgColor(string $color)

#### Apply Text Styles
* applyTextAlign(string $textAlign, ?string $verticalAlign = null)
* applyVerticalAlign(string $verticalAlign)
* applyTextCenter()
* applyTextWrap(bool $textWrap)
* applyTextRotation(int $degrees) (thanks to @jarrod-colluco)

Returns to [README.md](/README.md)
