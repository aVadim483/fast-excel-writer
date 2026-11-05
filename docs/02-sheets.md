## FastExcelWriter – Sheets

### Create, Select, and Remove Sheet

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

$sheet->setTabColor('#ff0099');

$sheet->setStateHidden();
$sheet->setStateVeryHidden();
$sheet->setStateVisible();
// OR
$sheet->setSheetState('hidden'); // or 'visible', 'veryHidden'
```


### Freeze Panes and Autofilter

```php
$sheet->setFreeze('B2');
$sheet->setAutofilter(1);
```

### Row's settings

You can set row options (styles and height) by the second argument when you use the function ```writeRow()```.
Note that in this case these styles will only be applied to those cells in the row where data is written

```php
// Write row data and set height
$rowStyle = [
    'fill-color' => '#fffeee',
    'border' => 'thin',
    'height' => 28,
];
$sheet->writeRow(['aaa', 'bbb', 'ccc'], $rowStyle);
```
Other way with the same result

```php
$sheet->writeRow(['aaa', 'bbb', 'ccc', null, 'eee'])
    ->applyFillColor('#fffeee')
    ->applyBorder('thin')
    ->applyRowHeight(28);

```
You can set row's height or visibility

```php
// Set height of row 2 to 33
$sheet->setRowHeight(2, 33);

// Set height of rows 3,5 and 7 to 33
$sheet->setRowHeight([3, 5, 7], 33);

// Set heights of several rows
$sheet->setRowHeights([1 => 20, 2 => 33, 3 => 40]);

// Hide row 8
$sheet->setRowVisible(8, false);

// Other way
$sheet->setRowHidden(8);

// Hide rows 9, 10, 11
$sheet->setRowVisible([9, 10, 11], false);

// Show row 10
$sheet->setRowVisible(10, true);
```
IMPORTANT: You can only use the setRowXX() functions on rows numbered at least as high as the current one.
See [Writing Row by Row vs Direct](03-writing.md#writing-row-by-row-vs-direct)
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

You can set the style for the entire sheet row or only for those cells in the row where data is written.
```php
// Style are applied to the entire sheet row
$sheet->setRowStyle(3, ['height' => 20]);
$sheet->setRowStyle('2:5', ['font-color' => '#f00']);
$sheet->setRowStyleArray([3 => $style1, 5 => $style2]);

// Set the style only for non-empty cells in a row
$sheet->setRowDataStyle(3, ['height' => 20]);
$sheet->setRowDataStyle('2:5', ['font-color' => '#f00']);
$sheet->setRowDataStyleArray([3 => $style1, 5 => $style2]);
```

### Column's settings

Column widths can be set in several ways

```php
// Set width of column D to 24
$sheet->setColWidth('D', 24);
$sheet->setColDataStyle('D', ['width' => 24]);

// Set width of specific columns
$sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);
// Set width of columns from 'A'
$sheet->setColWidths([10, 20, 30, 40], 24);

$colStyle = [
    'B' => ['width' => 10], 
    'C' => ['width' => 'auto'], 
    'E' => ['width' => 30], 
    'F' => ['width' => 40],
];
$sheet->setColDataStyleArray($colStyle);

```
You can define a minimal width of columns. Note that the minimum value has higher priority
```php
// Set minimum width to 20 
$sheet->setColMinWidth('D', 20);

// The value 10 will not be set because it is less than the minimum value
$sheet->setColWidth('D', 10);

// But width 30 will be set
$sheet->setColWidth('D', 30);

// The column width will be set to the width of the content, but not less than 20
$sheet->setColWidthAuto('D');

// Hide column B
$sheet->setColVisible('B', false);

// Other way
$sheet->setColHidden('B');

// Hide columns B, E, H
$sheet->setColVisible(['B', 'E', 'H'], false);

// Show column E
$sheet->setColVisible('E', true);
```

You can set the style for the entire sheet column or only for those cells in the column where data is written.
```php
// Style are applied to the entire sheet column
$sheet->setColStyle('B', $style); // style for cells of column 'B'
$sheet->setColStyle(2, $style); // column 'B' has a number 2
$sheet->setColStyle('C:F', $style); // style for range of columns
$sheet->setColStyle(['A', 'B', 'C'], $style); // options for several columns 'A', 'B' and 'C'
$sheet->setColStyleArray(['B' => ['width' => 20], 'C' => ['font-color' => '#f00']]);

// Set the style only for non-empty cells in a column
$sheet->setColDataStyle('B', ['width' => 20]);
$sheet->setColDataStyle(2, ['width' => 20]);
$sheet->setColDataStyle('B:D', ['width' => 'auto']);
$sheet->setColDataStyle(['A', 'B', 'C'], $style);
$sheet->setColDataStyleArray(['B' => $style1, 'C' => $style2]);
```

### Automatic column widths

```php
// Set auto width
$sheet->setColWidth('D', 'auto');
$sheet->setColWidthAuto('D');
$sheet->setColDataStyle('D', ['width' => 'auto']);

// Set width of specific columns
$sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);

$colStyle = [
    'B' => ['width' => 10], 
    'C' => ['width' => 'auto'], 
    'E' => ['width' => 30], 
    'F' => ['width' => 40],
];
$sheet->setColDataStyleArray($colStyle);
```

**IMPORTANT!** Automatic calculation of column width is a very complex task. When using this option, keep in mind the following:

1. The calculation is performed for each cell in the column as the sheet is filled. 
Therefore, automatic adjustment of the column width must be performed before you start writing data in the cells. 
If this adjustment is performed at the end of document creation, it has no effect.

2. Widths use character advance tables for Calibri, Arial, Tahoma and Times New Roman at 11 pt, scaled for size and selected for normal/bold/italic faces. Tables cover ASCII and Russian Cyrillic. Unknown fonts (including Aptos), other characters and complex text shaping use approximations. No GD extension or installed font files are required at runtime.

Content pixels are converted to Excel units using the maximum digit width of the workbook base font, with cell padding added once. The regression fixture `tests/auto-width-excel.json` contains 41 AutoFit cases from Excel 16.0, with actual saved XML widths. Its acceptance tolerance is the larger of 1.15 width units and 5%; it covers these fixtures, not every font, size or writing system. See Microsoft's [column width definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column) and [number format guidelines](https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format).

3. The library cannot calculate formulas. A formula without a supplied result is ignored; a formula with a known result uses that result when calculating the width.

Automatic widths handle zero, boolean values and RichText, including fragment fonts, sizes and wrapped lines. Numeric cells are measured using their displayed value rather than the longer of the raw and formatted values. Supported masks include decimal places, leading zeros, grouping/scaling, percentages, currencies, scientific notation, sections/conditions, dates, time and elapsed `[h]:mm`. Text cells retain their text, including leading zeros. English, German, French and Russian numeric separators are selected from the workbook locale; explicit supported locale codes in a mask take precedence.

Unsupported masks (for example fractions) fall back to the original value. Repeated `*` fill is ignored and `_` spacing is approximated by a space. A callback can supply display text for custom formats without changing the saved cell:

```php
$excel->setAutoWidthFormatter(function ($value, $format, $type, $locale) {
    return $format === '# ?/?' ? 'custom fraction text' : null;
});
```

The callback must return a string, or `null` to use the built-in formatter. It runs for each measured scalar cell and is not cached. It is also available as the `auto_width_formatter` workbook option and `Options::autoWidthFormatter()`.

Ordinary numeric date/time masks, such as `yyyy-mm-dd hh:mm:ss`, use a cached width sample (`0000-00-00 00:00:00`) instead of formatting each date. Single-digit fields are estimated at their maximum two-digit width. Month/day names, AM/PM, elapsed times and conditional sections retain value-based formatting. A supplied callback runs before this optimization and still receives the actual value.

Width results use bounded FIFO caches (4096 entries, up to 1 MiB of keys, skipping keys over 1 KiB). Parsed formats have a separate bounded cache. Workbook caches are released with the workbook; shared strings and the style registry have their own memory requirements. `Font::setFontWidthFactor()` affects subsequent measurements, including cached ones, without remeasuring already written rows.

Measurement cache keys reuse equal-width digit patterns: the supported font tables use tabular digits, so `Row 12345` and `Row 98765` can share a measurement. This applies after display formatting and does not alter stored values. Short Unicode text uses multibyte splitting; long text covered by the tables uses a native character-count path with the same advances. Other characters retain the existing fallback.

Native PHP integers with `General` or `@` use a direct width path without number-format parsing or locale conversion. The simple `0` mask uses this path for integers of up to 15 digits; larger formatted values retain the standard formatter. A bounded cache stores signed decimal lengths for one active font, with base-font conversion and the current width factor applied on retrieval. Custom masks and callbacks retain their usual behavior.

Native booleans use the same bounded font context with two literal measurements, `TRUE` and `FALSE`, independently of the number mask. Both boolean values participate in automatic widths. Font, workbook base-font and width-factor changes remain effective; a custom callback always takes precedence over this path.

Initial implementation measurements before the date/text optimizations, on Windows with Xdebug disabled (2026-11-05, median of five separate runs, baseline `5996193`):

| Scenario | PHP 7.4 before / after | PHP 8.3 before / after |
| --- | --- | --- |
| 100,000 unique static width calculations with four masks, retained process memory | 12.60 / 1.36 MiB | 12.58 / 0.96 MiB |
| Export 10,000 numeric rows, automatic widths enabled | 0.138 / 0.157 s | 0.133 / 0.146 s |
| Export 10,000 numeric rows, automatic widths disabled | 0.119 / 0.122 s | 0.109 / 0.112 s |

The more detailed measurement costs time when automatic widths are enabled. In a separate streaming run with a fixed style set and inline strings, memory stayed near 2.2 MiB (PHP 7.4) / 1.3 MiB (PHP 8.3) above the initial process usage from 100,000 through 1,000,000 rows across two sheets. These figures describe the measured workloads; shared strings and unique styles can still grow. GD/FreeType was evaluated and remains a separate possible extension because it requires font files and costs substantially more per uncached measurement.

After optimizing dates and text, the same 10,000-row export workloads were compared against the initial implementation (`3e8e6e1`), alternating versions over five separate runs:

| Scenario | PHP 7.4 reduction in export time | PHP 8.3 reduction in export time |
| --- | --- | --- |
| Numeric date/time mask | 44% | 36% |
| Text with the same prefix and varying digits | 16% | 20% |
| Unique text without digits | 3% | 11% |
| Long unique text | 33% | 32% |

Short repeated text did not show a consistent speedup. In the same comparison, numeric dates and text with varying digits were within about 3% of the original `5996193` timings; other workloads can still cost more than the original coarse font estimation. These are workload-specific results, not a speed guarantee for every string or format.

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
$sheet->setActiveCell('B2');

// Selecting a range of cells
$sheet->setActiveCell('B2:C3');
```

### Page Setup

```php
// Set page size
$sheet->pagePaperSizeA4();
$sheet->pagePaperSizeA3();
$sheet->pagePaperSizeLetter();
$sheet->pagePaperSizeLegal();

// Other sizes are defined in Excel::PAPERSIZE_* constants
$sheet->pagePaperSize(Excel::PAPERSIZE_A4);
$sheet->pagePaperSize(Excel::PAPERSIZE_B5);
$sheet->pagePaperSize(Excel::PAPERSIZE_JAPANESE_POSTCARD_ROTATED);

$sheet->pagePaperHeight('297mm');
$sheet->pagePaperWidth('21cm');

// Scaling the page
$sheet->pageScale(100); // 100%
$sheet->pageFitToWidth(1); // fit width to 1 page
$sheet->pageFitToHeight(0);// autofit height

// Set page orientation
$sheet->pagePortrait();
$sheet->pageLandscape();

// Set margins
$sheet->pageMargins([
        'left' => '0.5',
        'right' => '0.5',
        'top' => '1.0',
        'bottom' => '1.0',
        'header' => '0.5',
        'footer' => '0.5',
    ]);
// the same action    
$sheet
    ->pageMarginLeft(0.5)
    ->pageMarginRight(0.5)
    ->pageMarginTop(1.0)
    ->pageMarginBottom(1.0)
    ->pageMarginHeader(0.5)
    ->pageMarginFooter(0.5);

```
By default, the values are set in inches, 1 inch is 2.54 cm. So when you specify numeric values, they are specified in inches.

But you can also specify these values in centimeters or millimeters.

```php
$sheet->pageMarginLeft(0.5); // set left margin 0.5 inch
$sheet->pageMarginLeft('0.5cm'); // set left margin 0.5 centimeters
$sheet->pageMarginLeft('0.5mm'); // set left margin 0.5 millimeters
```

### Print settings

Specify printing area

```php
$sheet->setPrintArea('A2:F3');
$sheet->setPrintArea('A8:F10');

// Or several print areas in one call
$sheet->setPrintArea('A2:F3,A8:F10');
```

To repeat specific rows/columns at top/left of a printing page, use the following code:

```php
$sheet->setPrintTopRows('1')->setPrintLeftColumns('A');
```

The following code is an example of how to repeat row 1 to 5 and columns A-C on each printed page:

```php
$sheet->setPrintTopRows('1:5');
$sheet->setPrintLeftColumns('A:C');
```

To show/hide gridlines when printing, use the following code:

```php
$sheet->setPrintGridlines(true);
```

You can center print area on a page

```php
// horizontal centered (default param is true)
$sheet->setPrintHorizontalCentered();

// vertical centered
$sheet->setPrintVerticalCentered();

// centered both directions
$sheet->setPrintCentered();
```
### Print Header and Footer

```php
// Set the same Header and Footer for all Pages
$sheet->pageHeaderFooter('Print Header', 'Print Footer');

$sheet->pageHeader('Header for all pages');
$sheet->pageHeaderFirst('Header for the First page');
$sheet->pageHeaderOdd('Header for Odd pages');
$sheet->pageHeaderEven('Header for Even pages');

$sheet->pageFooter('Footer for all pages');
$sheet->pageFooterFirst('Footer for the First page');
$sheet->pageFooterOdd('Footer for Odd pages');
$sheet->pageFooterEven('Footer for Even pages');
```

When defining headers and footers, you can use special formatting codes that begin with `&`.  
Below is a complete, practical list.

#### Page and document fields

- **`&P`** — current page number
- **`&N`** — total number of pages
- **`&D`** — current date
- **`&T`** — current time
- **`&A`** — worksheet (sheet) name
- **`&F`** — file name
- **`&Z`** — file path
- **`&G`** — insert an image (picture in the header/footer)

```php
$sheet->pageFooter('Page &P of &N');
```

#### Font and text formatting

- **`&"FontName,Style"`** — set font and style
- **`&12`** — set font size (number = size)
```php
$sheet->pageFooter('&"Arial,Bold Italic"Page &P of &N');
```

Style toggles:

- **`&B`** — bold
- **`&I`** — italic
- **`&U`** — underline
- **`&E`** — double underline
- **`&S`** — strikethrough
- **`&X`** — superscript
- **`&Y`** — subscript

Color:

- **`&Krrggbb`** — text color (hex RGB)

#### Position within the header/footer

These define where text is aligned inside a single header/footer:

- **`&L`** — left section
- **`&C`** — center section
- **`&R`** — right section

```php
$sheet->pageHeaderFirst('&LLeft &CCenter &RRight');
```

#### Additional rules

- **`&&`** — prints a literal `&` (otherwise it is treated as a control code)

You can combine codes freely

