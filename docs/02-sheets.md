## FastExcelWriter - Sheets

### Create, Select and Remove Sheet

```php
// Create workbook with three named sheets 
$excel = Excel::create(['Jan', 'Feb', 'Mar']);

// Get the first sheet;
$sheet = $excel->getSheet();

// Get the sheet 'Jan';
$sheet = $excel->getSheet('Jan');

// Get the third sheet (named 'Mar');
$sheet = $excel->getSheet(3);

// make new sheet with name 'Total'
$sheet = $excel->makeSheet('Total');

$sheet->setName($sheetName);

// Removes the specified sheet
$excel->removeSheet('Total');
```

### Sheet Settings

```php
$sheet1->setDefaultFont($font);
$sheet1->setDefaultFontName($fontName);
$sheet1->setDefaultFontSize($fontSize);
$sheet1->setDefaultFontStyle($fontStyle);
$sheet1->setDefaultFontStyleBold();
$sheet1->setDefaultFontStyleItalic();
$sheet1->setDefaultFontStyleUnderline(true);
$sheet1->setDefaultFontStyleStrikethrough();
$sheet1->setDefaultFontColor($font);

```

### Page Settings

```php
$sheet1->pagePortrait();
$sheet1->pageLandscape();
$sheet1->pageFitToWidth(1); // fit width to 1 page
$sheet1->pageFitToHeight(1);// fit height to 1 page

$sheet1->pageMargins([
        'left' => '0.5',
        'right' => '0.5',
        'top' => '1.0',
        'bottom' => '1.0',
        'header' => '0.5',
        'footer' => '0.5',
    ]);
$sheet1->pageMarginLeft(0.5);
$sheet1->pageMarginRight(0.5);
$sheet1->pageMarginTop(1.0);
$sheet1->pageMarginBottom(1.0);
$sheet1->pageMarginHeader(0.5);
$sheet1->pageMarginFooter(0.5);

$sheet1->pagePaperSize(Excel::PAPERSIZE_A3);
$sheet1->pagePaperHeight('297mm');
$sheet1->pagePaperWidth('21cm');
```

### Freeze Panes and Autofilter

```php
$sheet1->setFreeze('B2');
$sheet1->setAutofilter(1);
```

### Row's settings

You can set row options (styles and height) by the second argument when you use the function ```writeRow()```.
Note that in this case these styles will only be applied to those cells in the row where data is written

```php
// Write row data and set height
$rowOptions = [
    'fill-color' => '#fffeee',
    'border' => 'thin',
    'height' => 28,
];
$sheet1->writeRow(['aaa', 'bbb', 'ccc'], $rowOptions);
```
Other way with the same result

```php
$sheet1->writeRow(['aaa', 'bbb', 'ccc', null, 'eee'])
    ->applyFillColor('#fffeee')
    ->applyBorder('thin')
    ->applyRowHeight(28);

```
You can set row's height or visibility

```php
// Set height of row 2 to 33
$sheet1->setRowHeight(2, 33);

// Set height of rows 3,5 and 7 to 33
$sheet1->setRowHeight([3, 5, 7], 33);

// Set heights of several rows
$sheet1->setRowHeights([1 => 20, 2 => 33, 3 => 40]);

// Hide row 8
$sheet1->setRowVisible(8, false);

// Other way
$sheet1->setRowHidden(8);

// Hide rows 9, 10, 11
$sheet1->setRowVisible([9, 10, 11], false);

// Show row 10
$sheet1->setRowVisible(10, true);
```
IMPORTANT: You can only use the setRowXX() functions on rows numbered at least as high as the current one.
See [Writing Row by Row vs Direct](/docs/03-writing.md#writing-row-by-row-vs-direct)
Therefore, the following code will throw an error "Row number must be greater then written rows"

```php
$sheet = $excel->sheet();
// Write row 1
$sheet->writeRow(['aaa1', 'bbb1', 'ccc1']);
// Write row 2
$sheet->writeRow(['aaa2', 'bbb2', 'ccc2']);
// Try to set height of previous row 1
$sheet->setRowHeight(1, 33);

```

### Column's settings

Column widths can be set in several ways

```php
// Set width of column D to 24
$this->setColWidth('D', 24);
$this->setColOptions('D', ['width' => 24]);
// Set auto width
$this->setColWidth('D', 'auto');
$this->setColWidthAuto('D');
$this->setColOptions('D', ['width' => 'auto']);

// Set width of specific columns
$sheet1->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);
// Set width of columns from 'A'
$sheet1->setColWidths([10, 20, 30, 40], 24);

$colOptions = [
    'B' => ['width' => 10], 
    'C' => ['width' => 'auto'], 
    'E' => ['width' => 30], 
    'F' => ['width' => 40],
];
$sheet1->setColOptions($colOptions);

```
You can define a minimal width of columns. Note that the minimum value has higher priority
```php
// Set minimum width to 20 
$this->setColMinWidth('D', 20);
// The value 10 will not be set because it is less than the minimum value
$this->setColWidth('D', 10);
// But width 30 will be set
$this->setColWidth('D', 30);
// The column width will be set to the width of the content, but not less than 20
$this->setColWidthAuto('D');
```

### Define Named Ranges

FastExcelWriter supports _named ranges_ and does not support _named formulae_.
A _named ranges_ provides a name reference to a cell or a range of cells.
All _named ranges_ are added to the workbook so all names must be unique, but you can define _named ranges_ in a sheet or in a workbook.

Also range names must start with a letter or underscore, have no spaces, and be no longer than 255 characters.

```php
$excel = Excel::create();
$excel->setFileName($outFileName);
$sheet = $excel->sheet();

// Named a single cell
$sheet->addNamedRange('B2', 'cell_name');

// Named range in a sheet
$sheet->addNamedRange('c2:e3', 'range_name');

// Add named range in a workbook (sheet name required)
$excel->addNamedRange('Sheet1!A1:F5', 'A1_F5');

// You can define name using applyNamedRange()
$sheet->writeCell(1000)->applyNamedRange('Value');
$sheet->writeCell(0.12)->applyNamedRange('Rate');
// Add the formula using names
$sheet->writeCell('=Value*Rate');

```

### Setting Active Cells

```php
// Selecting one active cell
$sheet1->setActiveCell('B2');

// Selecting a range of cells
$sheet1->setActiveCell('B2:C3');
```

Returns to [README.md](/README.md)
