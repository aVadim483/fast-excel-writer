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
$sheet->setDefaultFont($font);
$sheet->setDefaultFontName($fontName);
$sheet->setDefaultFontSize($fontSize);
$sheet->setDefaultFontStyle($fontStyle);
$sheet->setDefaultFontStyleBold();
$sheet->setDefaultFontStyleItalic();
$sheet->setDefaultFontStyleUnderline(true);
$sheet->setDefaultFontStyleStrikethrough();
$sheet->setDefaultFontColor($font);

```

### Page Settings

```php
$sheet->pagePortrait();
$sheet->pageLandscape();
$sheet->pageFitToWidth(1); // fit width to 1 page
$sheet->pageFitToHeight(1);// fit height to 1 page

$sheet->pageMargins([
        'left' => '0.5',
        'right' => '0.5',
        'top' => '1.0',
        'bottom' => '1.0',
        'header' => '0.5',
        'footer' => '0.5',
    ]);
$sheet->pageMarginLeft(0.5);
$sheet->pageMarginRight(0.5);
$sheet->pageMarginTop(1.0);
$sheet->pageMarginBottom(1.0);
$sheet->pageMarginHeader(0.5);
$sheet->pageMarginFooter(0.5);

$sheet->pagePaperSize(Excel::PAPERSIZE_A3);
$sheet->pagePaperHeight('297mm');
$sheet->pagePaperWidth('21cm');
```

### Freeze Panes and Autofilter

```php
$sheet->setFreeze('B2');
$sheet->setAutofilter(1);
```

### Setting Active Cells

```php
// Selecting one active cell
$sheet->setActiveCell('B2');

// Selecting a range of cells
$sheet->setActiveCell('B2:C3');
```

Returns to [README.md](/README.md)
