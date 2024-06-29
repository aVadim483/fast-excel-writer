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
Therefore, the following code will throw an error "Row number must be greater than written rows"

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
$sheet->setColWidth('D', 24);
$sheet->setColOptions('D', ['width' => 24]);
// Set auto width
$sheet->setColWidth('D', 'auto');
$sheet->setColWidthAuto('D');
$sheet->setColOptions('D', ['width' => 'auto']);

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
### Group/outline rows and columns

Set group level for the specified rows

```php
$sheet = $excel->sheet();

// the first level
$sheet->writeRow($rowData1)->applyRowOutlineLevel(1);
$sheet->writeRow($rowData2)->applyRowOutlineLevel(1);

// the second level
$sheet->writeRow($rowData3)->applyRowOutlineLevel(2);
$sheet->writeRow($rowData4)->applyRowOutlineLevel(2);

// back to the first level
$sheet->writeRow($rowData5)->applyRowOutlineLevel(1);

// write rows without grouping
$sheet->writeRow($rowData6);
$sheet->writeRow($rowData7);
```

You can set up grouping for future rows.

```php
// set level 1 for row 4
$sheet->setRowOutlineLevel(4, 1);

// set level 1 for rows 5, 6, 7
$sheet->setRowOutlineLevel([5, 6, 7], 1);

// set level 1 for rows from 9 to 15
$sheet->setRowOutlineLevel('9:15', 1);
// set level 2 for rows from 11 to 13
$sheet->setRowOutlineLevel('11:13', 2);
```

You can set up grouping for a sequence of rows.

```php
$sheet = $excel->sheet();

// Writing rows without grouping
$sheet->writeRow([...]);
$sheet->writeRow([...]);

// Increase group level (set level to 1)
$sheet->beginOutlineLevel();
$sheet->writeRow([...]);
$sheet->writeRow([...]);

// Increase group level again (set level to 2) with collapsing
$sheet->beginOutlineLevel(true);
$sheet->writeRow([...]);
$sheet
    ->writeCell('...')
    ->writeCell('...')
    ->writeCell('...')
    ->nextRow();
$sheet->writeRow([...]);

// Decrease group level (back to 1)
$sheet->endOutlineLevel();
$sheet->writeRow([...]);

// Set zero level
$sheet->endOutlineLevel();
```

Set group level for the specified columns

```php
$sheet->setColOutlineLevel('B', 1);
$sheet->setColOutlineLevel('C', 1);
$sheet->setColOutlineLevel('D', 1);

$sheet->setColOutlineLevel(['F', 'g', 'h', 'i', 'J'], 1);
$sheet->setColOutlineLevel('g:i', 2);

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

### Setting Active Sheet and Cells

You can select active (default) sheet in workbook

```php
// Set active (default) sheet by case-insensitive name
$excel->setActiveSheet($name);
```

To select active cell on specified sheet use the following code:

```php
// Selecting one active cell
$sheet1->setActiveCell('B2');

// Selecting a range of cells
$sheet1->setActiveCell('B2:C3');
```

### Print settings

Specify printing area

```php
$sheet->setPrintArea('A2:F3,A8:F10');
```

To repeat specific rows/columns at top/left of a printing page, use the following code:

```php
$sheet->setPrintTopRows('1')->setPrintLeftColumns('A');
```

The following code is an example of how to repeat row 1 to 5 on each printed page:

```php
$sheet->setPrintTopRows('1:5');
```

To show/hide gridlines when printing, use the following code:

```php
$sheet->setPrintGridlines(true);
```

Returns to [README.md](/README.md)
