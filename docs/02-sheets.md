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

### Setting Active Cells

```php
// Selecting one active cell
$sheet1->setActiveCell('B2');

// Selecting a range of cells
$sheet1->setActiveCell('B2:C3');
```

Returns to [README.md](/README.md)
