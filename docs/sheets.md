# FastExcelWriter

## Sheets

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

```

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

pagePortrait()
pageLandscape()
pageFitToWidth(1)  // fit width to 1 page
pageFitToHeight(1);// fit height to 1 page
->setFreeze('B2')
->setAutofilter(1);
