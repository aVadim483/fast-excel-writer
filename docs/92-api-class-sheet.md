# Class \avadim\FastExcelWriter\Sheet

---

* [__construct()](#__construct) – Sheet constructor
* [addCellStyle()](#addcellstyle) – Add additional styles to a cell
* [addChart()](#addchart) – Add a chart object to the specified range of cells
* [addConditionalFormatting()](#addconditionalformatting) – Add a conditional formatting object to the specified range of cells
* [addDataValidation()](#adddatavalidation) – Add a data validation object to the specified range of cells
* [addImage()](#addimage) – Add an image to the sheet from a local file, URL or image string in base64
* [addNamedRange()](#addnamedrange) – Define named range
* [addNote()](#addnote) – Add a note to the sheet
* [addStyle()](#addstyle) – Alias for 'addCellStyle()'
* [allowAutoFilter()](#allowautofilter) – AutoFilters should be allowed to operate when the sheet is protected
* [allowDeleteColumns()](#allowdeletecolumns) – Deleting columns should be allowed when the sheet is protected
* [allowDeleteRows()](#allowdeleterows) – Deleting rows should be allowed when the sheet is protected
* [allowEditObjects()](#alloweditobjects) – Objects are allowed to be edited when the sheet is protected
* [allowEditScenarios()](#alloweditscenarios) – Scenarios are allowed to be edited when the sheet is protected
* [allowFormatCells()](#allowformatcells) – Formatting cells should be allowed when the sheet is protected
* [allowFormatColumns()](#allowformatcolumns) – Formatting columns should be allowed when the sheet is protected
* [allowFormatRows()](#allowformatrows) – Formatting rows should be allowed when the sheet is protected
* [allowInsertColumns()](#allowinsertcolumns) – Inserting columns should be allowed when the sheet is protected
* [allowInsertHyperlinks()](#allowinserthyperlinks) – Inserting hyperlinks should be allowed when the sheet is protected
* [allowInsertRows()](#allowinsertrows) – Inserting rows should be allowed when the sheet is protected
* [allowPivotTables()](#allowpivottables) – PivotTables should be allowed to operate when the sheet is protected
* [allowSelectCells()](#allowselectcells) – Selection of any cells should be allowed when the sheet is protected
* [allowSelectLockedCells()](#allowselectlockedcells) – Selection of locked cells should be allowed when the sheet is protected
* [allowSelectUnlockedCells()](#allowselectunlockedcells) – Selection of unlocked cells should be allowed when the sheet is protected
* [allowSort()](#allowsort) – Sorting should be allowed when the sheet is protected
* [applyAlignLeft()](#applyalignleft) – Apply left alignment to content
* [applyAlignRight()](#applyalignright) – Apply right alignment to content
* [applyBgColor()](#applybgcolor) – Alias of 'applyFillColor()'
* [applyBorder()](#applyborder) – Sets all borders style
* [applyBorderBottom()](#applyborderbottom) – Apply bottom border style and color to the selected area
* [applyBorderLeft()](#applyborderleft) – Apply left border style and color to the selected area
* [applyBorderRight()](#applyborderright) – Apply right border style and color to the selected area
* [applyBorderTop()](#applybordertop) – Apply top border style and color to the selected area
* [applyColor()](#applycolor) – Alias of 'setFontColor()'
* [applyConditionalFormatting()](#applyconditionalformatting) – Apply conditional formatting to the selected area
* [applyDataValidation()](#applydatavalidation) – Apply data validation to the selected area
* [applyFillColor()](#applyfillcolor) – Fill background color
* [applyFillGradient()](#applyfillgradient) – Fill background by gradient
* [applyFont()](#applyfont) – Apply font settings to the selected area
* [applyFontColor()](#applyfontcolor) – Apply font color to the selected area
* [applyFontName()](#applyfontname) – Apply font name to the selected area
* [applyFontSize()](#applyfontsize) – Apply font size to the selected area
* [applyFontStyle()](#applyfontstyle) – Apply font style (bold, italic, etc.) to the selected area
* [applyFontStyleBold()](#applyfontstylebold) – Apply bold font style to the selected area
* [applyFontStyleItalic()](#applyfontstyleitalic) – Apply italic font style to the selected area
* [applyFontStyleStrikethrough()](#applyfontstylestrikethrough) – Apply strikethrough font style to the selected area
* [applyFontStyleUnderline()](#applyfontstyleunderline) – Apply underline font style to the selected area
* [applyFormat()](#applyformat) – Set value format for the selected area
* [applyHide()](#applyhide) – Set hidden protection for the selected area
* [applyIndentDistributed()](#applyindentdistributed) – Set distributed indent for the selected area
* [applyIndentLeft()](#applyindentleft) – Set left indent for the selected area
* [applyIndentRight()](#applyindentright) – Set right indent for the selected area
* [applyInnerBorder()](#applyinnerborder) – Apply inner border style and color to the selected area
* [applyNamedRange()](#applynamedrange) – Apply named range to the selected area
* [applyOuterBorder()](#applyouterborder) – Apply outer border style and color to the selected area
* [applyRowHeight()](#applyrowheight) – Sets height to the current row
* [applyRowOutlineLevel()](#applyrowoutlinelevel) – Set outline level for the current row
* [applyStyle()](#applystyle) – Apply the style
* [applyTextAlign()](#applytextalign) – Apply horizontal alignment to the selected area
* [applyTextCenter()](#applytextcenter) – Apply horizontal and vertical center alignment to the selected area
* [applyTextColor()](#applytextcolor) – Set text color for the selected area
* [applyTextRotation()](#applytextrotation) – Set text rotation for the selected area
* [applyTextWrap()](#applytextwrap) – Set text wrap for the selected area
* [applyUnlock()](#applyunlock) – Set unlock protection for the selected area
* [applyVerticalAlign()](#applyverticalalign) – Apply vertical alignment to the selected area
* [beginArea()](#beginarea) – Begin a new area
* [beginOutlineLevel()](#beginoutlinelevel) – Start a new outline level for rows
* [cell()](#cell) – Select a single cell or cell range in the current row
* [clearAreas()](#clearareas) – Clear all defined areas
* [endAreas()](#endareas)
* [endOutlineLevel()](#endoutlinelevel) – End the current outline level for rows
* [getCharts()](#getcharts) – Get all charts in the sheet
* [getColAttributes()](#getcolattributes) – Get columns attributes
* [getConditionalFormatting()](#getconditionalformatting) – Get all conditional formatting in the sheet
* [getCurrentCell()](#getcurrentcell) – Returns address of the current cell
* [getCurrentCol()](#getcurrentcol) – Returns current column letter
* [getCurrentColId()](#getcurrentcolid) – Get current column index (0-based)
* [getCurrentRow()](#getcurrentrow) – Returns current row number
* [getCurrentRowId()](#getcurrentrowid) – Get current row index (0-based)
* [getDataValidations()](#getdatavalidations) – Get all data validations in the sheet
* [getDefaultStyle()](#getdefaultstyle) – Returns default style
* [getHeaderFooterOptions()](#getheaderfooteroptions) – Get header and footer options
* [getHyperlinks()](#gethyperlinks) – Returns added hyperlinks
* [getImages()](#getimages) – Get all images in the sheet
* [getLastCell()](#getlastcell) – Get address of the last touched cell
* [getLastRange()](#getlastrange) – Get address of the last touched range
* [getMergedCells()](#getmergedcells) – Returns merged cells
* [getName()](#getname) – Get sheet name
* [getNamedRanges()](#getnamedranges) – Returns named ranges with full addresses
* [getNotes()](#getnotes) – Get all notes in the sheet
* [getOutlineLevel()](#getoutlinelevel) – Get the current outline level for rows
* [isName()](#isname) – Case-insensitive name checking
* [isRightToLeft()](#isrighttoleft) – Check if the sheet is right-to-left
* [makeArea()](#makearea) – Make area for writing
* [mergeCells()](#mergecells) – Merge cells
* [mergeRelCells()](#mergerelcells) – Merge relative cells
* [nextCell()](#nextcell) – Move pointer to the next cell
* [nextRow()](#nextrow) – Move to the next row
* [pageFitToHeight()](#pagefittoheight) – Set page to fit to height
* [pageFitToWidth()](#pagefittowidth) – Set page to fit to width
* [pageFooter()](#pagefooter) – Set the footer for all pages
* [pageFooterEven()](#pagefootereven) – Set footer for even pages
* [pageFooterFirst()](#pagefooterfirst) – Set the footer for the first page
* [pageFooterOdd()](#pagefooterodd) – Set footer for odd pages
* [pageHeader()](#pageheader) – Set page header for all pages
* [pageHeaderEven()](#pageheadereven) – Set header for even pages
* [pageHeaderFirst()](#pageheaderfirst) – Set the header for the first page
* [pageHeaderFooter()](#pageheaderfooter) – Set the center header and footer for all pages
* [pageHeaderOdd()](#pageheaderodd) – Set header for odd pages
* [pageLandscape()](#pagelandscape) – Set page orientation as Landscape
* [pageMarginBottom()](#pagemarginbottom) – Bottom Page Margin in mm|cm|in
* [pageMarginFooter()](#pagemarginfooter) – Footer Page Margin in mm|cm|in
* [pageMarginHeader()](#pagemarginheader) – Header Page Margin in mm|cm|in
* [pageMarginLeft()](#pagemarginleft) – Left Page Margin in mm|cm|in
* [pageMarginRight()](#pagemarginright) – Right page margin in mm|cm|in
* [pageMargins()](#pagemargins) – Page margins for a sheet or a custom sheet view in mm|cm|in
* [pageMarginTop()](#pagemargintop) – Top Page Margin in mm|cm|in
* [pageOrientationLandscape()](#pageorientationlandscape) – Set page orientation as Landscape, alias of pageLandscape()
* [pageOrientationPortrait()](#pageorientationportrait) – Set page orientation as Portrait, alias of pagePortrait()
* [pagePaperHeight()](#pagepaperheight) – Height of custom paper as a number followed by a unit identifier mm|cm|in (ex: 297mm, 11in)
* [pagePaperSize()](#pagepapersize) – Set Paper size (when paperHeight and paperWidth are specified, paperSize should be ignored)
* [pagePaperSizeA3()](#pagepapersizea3) – Set Paper Size to A3 (when paperHeight and paperWidth are specified, paperSize should be ignored)
* [pagePaperSizeA4()](#pagepapersizea4) – Set Paper Size to A4 (when paperHeight and paperWidth are specified, paperSize should be ignored)
* [pagePaperSizeLegal()](#pagepapersizelegal) – Set Paper Size to Legal (when paperHeight and paperWidth are specified, paperSize should be ignored)
* [pagePaperSizeLetter()](#pagepapersizeletter) – Set Paper Size to Letter (when paperHeight and paperWidth are specified, paperSize should be ignored)
* [pagePaperWidth()](#pagepaperwidth) – Width of custom paper as a number followed by a unit identifier mm|cm|in (ex: 21cm, 8.5in)
* [pagePortrait()](#pageportrait) – Set page orientation as Portrait
* [pageScale()](#pagescale) – Set page scale
* [protect()](#protect) – Protect sheet
* [setActiveCell()](#setactivecell) – Set active cell
* [setAutoFilter()](#setautofilter) – Set auto filter
* [setBgColor()](#setbgcolor) – Set background color for the specific cell or range
* [setBottomNodesOptions()](#setbottomnodesoptions) – Set multiple options for a bottom node
* [setCellStyle()](#setcellstyle) – Set style for the specific cell
* [setColAutoWidth()](#setcolautowidth) – Alias of setColWidthAuto($col)
* [setColDataStyle()](#setcoldatastyle) – Set styles of column cells (colors, formats, etc.)
* [setColDataStyleArray()](#setcoldatastylearray) – Set style of column cells (colors, formats, etc.)
* [setColFormat()](#setcolformat) – Set a format of single or multiple column(s)
* [setColFormats()](#setcolformats) – Set formats of columns
* [setColFormula()](#setcolformula) – Set formula for single or multiple column(s)
* [setColHidden()](#setcolhidden) – Set a column as hidden
* [setColMinWidth()](#setcolminwidth) – Setting a minimal column's width
* [setColMinWidths()](#setcolminwidths) – Setting a multiple column's minimal width
* [setColOutlineLevel()](#setcoloutlinelevel) – Set a column outline level
* [setColStyle()](#setcolstyle) – Set style of single or multiple column(s)
* [setColStyleArray()](#setcolstylearray) – Set styles of single or multiple column(s)
* [setColVisible()](#setcolvisible) – Show/hide a column
* [setColWidth()](#setcolwidth) – Set a width of single or multiple column(s)
* [setColWidthAuto()](#setcolwidthauto) – Set auto width of single or multiple column(s)
* [setColWidths()](#setcolwidths) – Setting a multiple column's width
* [setDefaultFont()](#setdefaultfont) – Set default font settings for the sheet
* [setDefaultFontColor()](#setdefaultfontcolor) – Set default font color
* [setDefaultFontName()](#setdefaultfontname) – Set default font name for the sheet
* [setDefaultFontSize()](#setdefaultfontsize) – Set default font size for the sheet
* [setDefaultFontStyle()](#setdefaultfontstyle) – Set default font style (bold, italic, etc.) for the sheet
* [setDefaultFontStyleBold()](#setdefaultfontstylebold) – Set default font style as bold for the sheet
* [setDefaultFontStyleItalic()](#setdefaultfontstyleitalic) – Set default font style as italic
* [setDefaultFontStyleStrikethrough()](#setdefaultfontstylestrikethrough) – Set default font style as strikethrough
* [setDefaultFontStyleUnderline()](#setdefaultfontstyleunderline) – Set default font style as underline
* [setDefaultStyle()](#setdefaultstyle) – Sets default style
* [setFormat()](#setformat) – Set value format for the specific cell or range
* [setFormula()](#setformula) – Set a formula to the single cell or to the cell range
* [setFreeze()](#setfreeze) – Freeze rows/columns
* [setFreezeColumns()](#setfreezecolumns) – Freeze columns
* [setFreezeRows()](#setfreezerows) – Freeze rows
* [setName()](#setname) – Set sheet name
* [setOuterBorder()](#setouterborder) – Set outer border for the specific range
* [setPrintArea()](#setprintarea) – Set print area for the sheet
* [setPrintCentered()](#setprintcentered) – Center the print area horizontally and vertically
* [setPrintGridlines()](#setprintgridlines) – Show grid lines in the print area
* [setPrintHorizontalCentered()](#setprinthorizontalcentered) – Center the print area horizontally
* [setPrintLeftColumns()](#setprintleftcolumns) – Set left columns to repeat on every printed page
* [setPrintRowAndColumnHeading()](#setprintrowandcolumnheading) – Print row and column headings in the printout
* [setPrintTitles()](#setprinttitles) – Set rows to repeat at top and columns to repeat at left when printing
* [setPrintTopRows()](#setprinttoprows) – Set top rows to repeat on every printed page
* [setPrintVerticalCentered()](#setprintverticalcentered) – Center the print area vertically
* [setRowDataStyle()](#setrowdatastyle) – Style are applied only to non-empty cells in a row (or row range)
* [setRowDataStyleArray()](#setrowdatastylearray) – Styles are applied only to non-empty cells in a rows
* [setRowHeight()](#setrowheight) – Height of a specific row
* [setRowHeights()](#setrowheights) – Multiple rows height
* [setRowHidden()](#setrowhidden) – Hide a specific row
* [setRowOutlineLevel()](#setrowoutlinelevel) – Set an outline level for a specific row or range of rows
* [setRowStyle()](#setrowstyle) – The style is applied to the entire sheet row (even if it is empty)
* [setRowStyleArray()](#setrowstylearray) – Styles are applied to the entire sheet row (even if it is empty)
* [setRowVisible()](#setrowvisible) – Hide/show a specific row
* [setShowGridLines()](#setshowgridlines) – Turn on/off grid lines
* [setStateHidden()](#setstatehidden) – Make the sheet hidden
* [setStateVeryHidden()](#setstateveryhidden) – Make the sheet very hidden
* [setStateVisible()](#setstatevisible) – Make the sheet visible
* [setStyle()](#setstyle) – Alias for 'setCellStyle()'
* [setTabColor()](#settabcolor) – Set color for the sheet tab
* [setTopLeftCell()](#settopleftcell) – Set the top left cell for writing
* [setValue()](#setvalue) – Set a value to the single cell or to the cell range
* [skipRow()](#skiprow) – Skip rows
* [unprotect()](#unprotect) – Unprotect sheet
* [withLastCell()](#withlastcell) – Select the last written cell for applying
* [withLastRow()](#withlastrow) – Select the last written row for applying
* [withRange()](#withrange) – Select a custom range for applying
* [writeArray()](#writearray) – Write values from a two-dimensional array (alias of writeRows)
* [writeArrayTo()](#writearrayto) – Write 2d array from the specified cell
* [writeCell()](#writecell) – Write value to the current cell and move a pointer to the next cell in the row
* [writeCells()](#writecells) – Write several values into cells of one row
* [writeHeader()](#writeheader) – Write a header row with optional styles and formats for columns
* [writeRow()](#writerow) – Write values to the current row
* [writeRows()](#writerows) – Write several rows from a two-dimensional array
* [writeTo()](#writeto) – Write value to the specified cell and move a pointer to the next cell in the row

---

## __construct()

---

```php
public function __construct(string $sheetName)
```
_Sheet constructor_

### Parameters

* `string $sheetName`

---

## addCellStyle()

---

```php
public function addCellStyle(string $cellAddr, $style): Sheet
```
_Add additional styles to a cell_

### Parameters

* `string $cellAddr`
* `array|Style $style`

---

## addChart()

---

```php
public function addChart(string $range, 
                         avadim\FastExcelWriter\Charts\Chart $chart): Sheet
```
_Add a chart object to the specified range of cells_

### Parameters

* `string $range` – Set the position where the chart should appear in the worksheet
* `Chart $chart` – Chart object

---

## addConditionalFormatting()

---

```php
public function addConditionalFormatting(string $range, $conditionals): Sheet
```
_Add a conditional formatting object to the specified range of cells_

### Parameters

* `string $range`
* `Conditional|Conditional[] $conditionals`

---

## addDataValidation()

---

```php
public function addDataValidation(string $range, 
                                  avadim\FastExcelWriter\DataValidation\DataValidation $validation): Sheet
```
_Add a data validation object to the specified range of cells_

### Parameters

* `string $range`
* `DataValidation $validation`

---

## addImage()

---

```php
public function addImage(string $cell, string $imageFile, 
                         ?array $imageStyle = []): Sheet
```
_Add an image to the sheet from a local file, URL or image string in base64_

### Parameters

* `string $cell` – Cell address
* `string $imageFile` – URL, local path or image string in base64
* `array|null $imageStyle` – \['width' => ..., 'height' => ..., 'hyperlink' => ...]

---

### Examples

```php
$sheet->addImage('A1', 'path/to/file');
$sheet->addImage('A1', 'path/to/file', ['width' => 100]);
```


---

## addNamedRange()

---

```php
public function addNamedRange(string $range, string $name): Sheet
```
_Define named range_

### Parameters

* `string $range`
* `string $name`

---

### Examples

```php
$sheet->addNamedRange('B3:C5', 'Demo');
```


---

## addNote()

---

```php
public function addNote($cell, $comment, array $noteStyle = []): Sheet
```
_Add a note to the sheet_

### Parameters

* `string|mixed $cell`
* `string|array|null $comment`
* `array $noteStyle`

---

### Examples

```php
$sheet->addNote('A1', $noteText, $noteStyle);
$sheet->writeCell($cellValue)->addNote($noteText, $noteStyle);
```


---

## addStyle()

---

```php
public function addStyle(string $cellAddr, $style): Sheet
```
_Alias for 'addCellStyle()'_

### Parameters

* `string $cellAddr`
* `array|Style $style`

---

## allowAutoFilter()

---

```php
public function allowAutoFilter(?bool $allow = true): Sheet
```
_AutoFilters should be allowed to operate when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowDeleteColumns()

---

```php
public function allowDeleteColumns(?bool $allow = true): Sheet
```
_Deleting columns should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowDeleteRows()

---

```php
public function allowDeleteRows(?bool $allow = true): Sheet
```
_Deleting rows should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowEditObjects()

---

```php
public function allowEditObjects(?bool $allow = true): Sheet
```
_Objects are allowed to be edited when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowEditScenarios()

---

```php
public function allowEditScenarios(?bool $allow = true): Sheet
```
_Scenarios are allowed to be edited when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowFormatCells()

---

```php
public function allowFormatCells(?bool $allow = true): Sheet
```
_Formatting cells should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowFormatColumns()

---

```php
public function allowFormatColumns(?bool $allow = true): Sheet
```
_Formatting columns should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowFormatRows()

---

```php
public function allowFormatRows(?bool $allow = true): Sheet
```
_Formatting rows should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowInsertColumns()

---

```php
public function allowInsertColumns(?bool $allow = true): Sheet
```
_Inserting columns should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowInsertHyperlinks()

---

```php
public function allowInsertHyperlinks(?bool $allow = true): Sheet
```
_Inserting hyperlinks should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowInsertRows()

---

```php
public function allowInsertRows(?bool $allow = true): Sheet
```
_Inserting rows should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowPivotTables()

---

```php
public function allowPivotTables(?bool $allow = true): Sheet
```
_PivotTables should be allowed to operate when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowSelectCells()

---

```php
public function allowSelectCells(?bool $allow = true): Sheet
```
_Selection of any cells should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowSelectLockedCells()

---

```php
public function allowSelectLockedCells(?bool $allow = true): Sheet
```
_Selection of locked cells should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowSelectUnlockedCells()

---

```php
public function allowSelectUnlockedCells(?bool $allow = true): Sheet
```
_Selection of unlocked cells should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## allowSort()

---

```php
public function allowSort(?bool $allow = true): Sheet
```
_Sorting should be allowed when the sheet is protected_

### Parameters

* `bool|null $allow`

---

## applyAlignLeft()

---

```php
public function applyAlignLeft(): Sheet
```
_Apply left alignment to content_

### Parameters

_None_

---

## applyAlignRight()

---

```php
public function applyAlignRight(): Sheet
```
_Apply right alignment to content_

### Parameters

_None_

---

## applyBgColor()

---

```php
public function applyBgColor(string $color, ?string $pattern = null): Sheet
```
_Alias of 'applyFillColor()'_

### Parameters

* `string $color`
* `string|null $pattern`

---

## applyBorder()

---

```php
public function applyBorder(string $style, ?string $color = '#000000'): Sheet
```
_Sets all borders style_

### Parameters

* `string $style`
* `string|null $color`

---

## applyBorderBottom()

---

```php
public function applyBorderBottom(string $style, 
                                  ?string $color = '#000000'): Sheet
```
_Apply bottom border style and color to the selected area_

### Parameters

* `string $style`
* `string|null $color`

---

## applyBorderLeft()

---

```php
public function applyBorderLeft(string $style, 
                                ?string $color = '#000000'): Sheet
```
_Apply left border style and color to the selected area_

### Parameters

* `string $style`
* `string|null $color`

---

## applyBorderRight()

---

```php
public function applyBorderRight(string $style, 
                                 ?string $color = '#000000'): Sheet
```
_Apply right border style and color to the selected area_

### Parameters

* `string $style`
* `string|null $color`

---

## applyBorderTop()

---

```php
public function applyBorderTop(string $style, 
                               ?string $color = '#000000'): Sheet
```
_Apply top border style and color to the selected area_

### Parameters

* `string $style`
* `string|null $color`

---

## applyColor()

---

```php
public function applyColor(string $color): Sheet
```
_Alias of 'setFontColor()'_

### Parameters

* `string $color`

---

## applyConditionalFormatting()

---

```php
public function applyConditionalFormatting($conditionals): Sheet
```
_Apply conditional formatting to the selected area_

### Parameters

* `Conditional|Conditional[] $conditionals`

---

## applyDataValidation()

---

```php
public function applyDataValidation(avadim\FastExcelWriter\DataValidation\DataValidation $validation): Sheet
```
_Apply data validation to the selected area_

### Parameters

* `DataValidation $validation`

---

## applyFillColor()

---

```php
public function applyFillColor(string $color, ?string $pattern = null): Sheet
```
_Fill background color_

### Parameters

* `string $color`
* `string|null $pattern`

---

## applyFillGradient()

---

```php
public function applyFillGradient(string $color1, string $color2, 
                                  ?int $degree = null): Sheet
```
_Fill background by gradient_

### Parameters

* `string $color1`
* `string $color2`
* `int|null $degree`

---

## applyFont()

---

```php
public function applyFont(string $fontName, ?int $fontSize = null, 
                          ?string $fontStyle = null, 
                          ?string $fontColor = null): Sheet
```
_Apply font settings to the selected area_

### Parameters

* `string $fontName`
* `int|null $fontSize`
* `string|null $fontStyle`
* `string|null $fontColor`

---

## applyFontColor()

---

```php
public function applyFontColor(string $fontColor): Sheet
```
_Apply font color to the selected area_

### Parameters

* `string $fontColor`

---

## applyFontName()

---

```php
public function applyFontName(string $fontName): Sheet
```
_Apply font name to the selected area_

### Parameters

* `string $fontName`

---

## applyFontSize()

---

```php
public function applyFontSize(float $fontSize): Sheet
```
_Apply font size to the selected area_

### Parameters

* `float $fontSize`

---

## applyFontStyle()

---

```php
public function applyFontStyle(string $fontStyle): Sheet
```
_Apply font style (bold, italic, etc.) to the selected area_

### Parameters

* `string $fontStyle`

---

## applyFontStyleBold()

---

```php
public function applyFontStyleBold(): Sheet
```
_Apply bold font style to the selected area_

### Parameters

_None_

---

## applyFontStyleItalic()

---

```php
public function applyFontStyleItalic(): Sheet
```
_Apply italic font style to the selected area_

### Parameters

_None_

---

## applyFontStyleStrikethrough()

---

```php
public function applyFontStyleStrikethrough(): Sheet
```
_Apply strikethrough font style to the selected area_

### Parameters

_None_

---

## applyFontStyleUnderline()

---

```php
public function applyFontStyleUnderline(?bool $double = false): Sheet
```
_Apply underline font style to the selected area_

### Parameters

* `bool|null $double`

---

## applyFormat()

---

```php
public function applyFormat($format): Sheet
```
_Set value format for the selected area_

### Parameters

* `string|array $format`

---

## applyHide()

---

```php
public function applyHide(?bool $hide = true): Sheet
```
_Set hidden protection for the selected area_

### Parameters

* `bool|null $hide`

---

## applyIndentDistributed()

---

```php
public function applyIndentDistributed(int $indent): Sheet
```
_Set distributed indent for the selected area_

### Parameters

* `int $indent`

---

## applyIndentLeft()

---

```php
public function applyIndentLeft(int $indent): Sheet
```
_Set left indent for the selected area_

### Parameters

* `int $indent`

---

## applyIndentRight()

---

```php
public function applyIndentRight(int $indent): Sheet
```
_Set right indent for the selected area_

### Parameters

* `int $indent`

---

## applyInnerBorder()

---

```php
public function applyInnerBorder(string $style, 
                                 ?string $color = '#000000'): Sheet
```
_Apply inner border style and color to the selected area_

### Parameters

* `string $style`
* `string|null $color`

---

## applyNamedRange()

---

```php
public function applyNamedRange(string $name): Sheet
```
_Apply named range to the selected area_

### Parameters

* `string $name`

---

## applyOuterBorder()

---

```php
public function applyOuterBorder(string $style, 
                                 ?string $color = '#000000'): Sheet
```
_Apply outer border style and color to the selected area_

### Parameters

* `string $style`
* `string|null $color`

---

## applyRowHeight()

---

```php
public function applyRowHeight(float $height): Sheet
```
_Sets height to the current row_

### Parameters

* `float $height`

---

## applyRowOutlineLevel()

---

```php
public function applyRowOutlineLevel(int $outlineLevel): Sheet
```
_Set outline level for the current row_

### Parameters

* `int $outlineLevel`

---

## applyStyle()

---

```php
public function applyStyle($style): Sheet
```
_Apply the style_

### Parameters

* `array|Style $style`

---

## applyTextAlign()

---

```php
public function applyTextAlign(string $textAlign, 
                               ?string $verticalAlign = null): Sheet
```
_Apply horizontal alignment to the selected area_

### Parameters

* `string $textAlign`
* `string|null $verticalAlign`

---

## applyTextCenter()

---

```php
public function applyTextCenter(): Sheet
```
_Apply horizontal and vertical center alignment to the selected area_

### Parameters

_None_

---

## applyTextColor()

---

```php
public function applyTextColor(string $color): Sheet
```
_Set text color for the selected area_

### Parameters

* `string $color`

---

## applyTextRotation()

---

```php
public function applyTextRotation(int $degrees): Sheet
```
_Set text rotation for the selected area_

### Parameters

* `int $degrees`

---

## applyTextWrap()

---

```php
public function applyTextWrap(?bool $textWrap = true): Sheet
```
_Set text wrap for the selected area_

### Parameters

* `bool|null $textWrap`

---

## applyUnlock()

---

```php
public function applyUnlock(?bool $unlock = true): Sheet
```
_Set unlock protection for the selected area_

### Parameters

* `bool $unlock`

---

## applyVerticalAlign()

---

```php
public function applyVerticalAlign(string $verticalAlign): Sheet
```
_Apply vertical alignment to the selected area_

### Parameters

* `string $verticalAlign`

---

## beginArea()

---

```php
public function beginArea(?string $cellAddress = null): Area
```
_Begin a new area_

### Parameters

* `string|null $cellAddress` – Upper left cell of area

---

## beginOutlineLevel()

---

```php
public function beginOutlineLevel(?bool $collapsed = false): Sheet
```
_Start a new outline level for rows_

### Parameters

* `bool|null $collapsed`

---

## cell()

---

```php
public function cell($cellAddress): Sheet
```
_Select a single cell or cell range in the current row_

### Parameters

* `string|array $cellAddress`

---

### Examples

```php
$sheet->cell('B5')->writeCell($value);
$sheet->cell('B5:C7')->applyBorder('thin');
$sheet->cell(['col' => 2, 'row' => 5])->applyUnlock();
$sheet->cell([2, 5])->applyColor($color);
```


---

## clearAreas()

---

```php
public function clearAreas(): Sheet
```
_Clear all defined areas_

### Parameters

_None_

---

## endAreas()

---

```php
public function endAreas(): Sheet
```


### Parameters

_None_

---

## endOutlineLevel()

---

```php
public function endOutlineLevel(): Sheet
```
_End the current outline level for rows_

### Parameters

_None_

---

## getCharts()

---

```php
public function getCharts(): array
```
_Get all charts in the sheet_

### Parameters

_None_

---

## getColAttributes()

---

```php
public function getColAttributes(): array
```
_Get columns attributes_

### Parameters

_None_

---

## getConditionalFormatting()

---

```php
public function getConditionalFormatting(): array
```
_Get all conditional formatting in the sheet_

### Parameters

_None_

---

## getCurrentCell()

---

```php
public function getCurrentCell(): string
```
_Returns address of the current cell_

### Parameters

_None_

---

## getCurrentCol()

---

```php
public function getCurrentCol(): string
```
_Returns current column letter_

### Parameters

_None_

---

## getCurrentColId()

---

```php
public function getCurrentColId(): int
```
_Get current column index (0-based)_

### Parameters

_None_

---

## getCurrentRow()

---

```php
public function getCurrentRow(): int
```
_Returns current row number_

### Parameters

_None_

---

## getCurrentRowId()

---

```php
public function getCurrentRowId(): int
```
_Get current row index (0-based)_

### Parameters

_None_

---

## getDataValidations()

---

```php
public function getDataValidations(): array
```
_Get all data validations in the sheet_

### Parameters

_None_

---

## getDefaultStyle()

---

```php
public function getDefaultStyle(): array
```
_Returns default style_

### Parameters

_None_

---

## getHeaderFooterOptions()

---

```php
public function getHeaderFooterOptions(): array
```
_Get header and footer options_

### Parameters

_None_

---

## getHyperlinks()

---

```php
public function getHyperlinks(): array
```
_Returns added hyperlinks_

### Parameters

_None_

---

## getImages()

---

```php
public function getImages(): array
```
_Get all images in the sheet_

### Parameters

_None_

---

## getLastCell()

---

```php
public function getLastCell(?bool $absolute = false): string
```
_Get address of the last touched cell_

### Parameters

* `bool|null $absolute`

---

## getLastRange()

---

```php
public function getLastRange(?bool $absolute = false): string
```
_Get address of the last touched range_

### Parameters

* `bool|null $absolute`

---

## getMergedCells()

---

```php
public function getMergedCells(): array
```
_Returns merged cells_

### Parameters

_None_

---

## getName()

---

```php
public function getName(): string
```
_Get sheet name_

### Parameters

_None_

---

## getNamedRanges()

---

```php
public function getNamedRanges(): array
```
_Returns named ranges with full addresses_

### Parameters

_None_

---

## getNotes()

---

```php
public function getNotes(): array
```
_Get all notes in the sheet_

### Parameters

_None_

---

## getOutlineLevel()

---

```php
public function getOutlineLevel(): int
```
_Get the current outline level for rows_

### Parameters

_None_

---

## isName()

---

```php
public function isName(string $name): bool
```
_Case-insensitive name checking_

### Parameters

* `string $name`

---

## isRightToLeft()

---

```php
public function isRightToLeft(): bool
```
_Check if the sheet is right-to-left_

### Parameters

_None_

---

## makeArea()

---

```php
public function makeArea(string $range): Area
```
_Make area for writing_

### Parameters

* `string $range` – A1:Z9 or R1C1:R9C28

---

## mergeCells()

---

```php
public function mergeCells($rangeSet, ?int $mergeFlag = 0): Sheet
```
_Merge cells_

### Parameters

* `array|string|int $rangeSet`
* `int|null $mergeFlag` – Action in case of intersection: 0 - exception; 1 - replace; 2 - keep; -1 - skip intersection check

---

### Examples

```php
$sheet->mergeCells('A1:C3');
$sheet->mergeCells(['A1:B2', 'C1:D2']);
$sheet->mergeCells('B5:C7', $value, Sheet:MERGE_NO_CHECK); // don't check for intersection of merged cells
```


---

## mergeRelCells()

---

```php
public function mergeRelCells($rangeSet): Sheet
```
_Merge relative cells_

### Parameters

* `array|string|int $rangeSet`

---

### Examples

```php
$sheet->mergeCells('C3:E8');
$sheet->mergeCells(3); // 3 columns of current row, equivalent of mergeCells('A5:C5') if current row is 5
$sheet->mergeCells(['RC3:RC5', 'RC6:RC7']); // equivalent of mergeCells(['C7:E7', 'F7:G7']) if current row is 7
```


---

## nextCell()

---

```php
public function nextCell(): Sheet
```
_Move pointer to the next cell_

### Parameters

_None_

---

## nextRow()

---

```php
public function nextRow($style, ?bool $forceRow = false): Sheet
```
_Move to the next row_

### Parameters

* `array|Style|null $style`
* `bool|null $forceRow`

---

## pageFitToHeight()

---

```php
public function pageFitToHeight($numPage): Sheet
```
_Set page to fit to height_

### Parameters

* `int|string|null $numPage`

---

## pageFitToWidth()

---

```php
public function pageFitToWidth($numPage): Sheet
```
_Set page to fit to width_

### Parameters

* `int|string|null $numPage`

---

## pageFooter()

---

```php
public function pageFooter($footer): Sheet
```
_Set the footer for all pages_

### Parameters

* `string|array $footer`

---

## pageFooterEven()

---

```php
public function pageFooterEven($footer): Sheet
```
_Set footer for even pages_

### Parameters

* `string|array $footer`

---

## pageFooterFirst()

---

```php
public function pageFooterFirst($footer): Sheet
```
_Set the footer for the first page_

### Parameters

* `string|array $footer`

---

## pageFooterOdd()

---

```php
public function pageFooterOdd(string $footer): Sheet
```
_Set footer for odd pages_

### Parameters

* `string $footer`

---

## pageHeader()

---

```php
public function pageHeader($header): Sheet
```
_Set page header for all pages_

### Parameters

* `string|array $header`

---

## pageHeaderEven()

---

```php
public function pageHeaderEven($header): Sheet
```
_Set header for even pages_

### Parameters

* `string|array $header`

---

## pageHeaderFirst()

---

```php
public function pageHeaderFirst($header): Sheet
```
_Set the header for the first page_

### Parameters

* `string|array $header`

---

## pageHeaderFooter()

---

```php
public function pageHeaderFooter(?string $header, ?string $footer): Sheet
```
_Set the center header and footer for all pages_

### Parameters

* `string|null $header`
* `string|null $footer`

---

## pageHeaderOdd()

---

```php
public function pageHeaderOdd($header): Sheet
```
_Set header for odd pages_

### Parameters

* `string|array $header`

---

## pageLandscape()

---

```php
public function pageLandscape(): Sheet
```
_Set page orientation as Landscape_

### Parameters

_None_

---

## pageMarginBottom()

---

```php
public function pageMarginBottom($value): Sheet
```
_Bottom Page Margin in mm|cm|in_

### Parameters

* `string|float $value`

---

## pageMarginFooter()

---

```php
public function pageMarginFooter($value): Sheet
```
_Footer Page Margin in mm|cm|in_

### Parameters

* `string|float $value`

---

## pageMarginHeader()

---

```php
public function pageMarginHeader($value): Sheet
```
_Header Page Margin in mm|cm|in_

### Parameters

* `string|float $value`

---

## pageMarginLeft()

---

```php
public function pageMarginLeft($value): Sheet
```
_Left Page Margin in mm|cm|in_

### Parameters

* `string|float $value`

---

## pageMarginRight()

---

```php
public function pageMarginRight($value): Sheet
```
_Right page margin in mm|cm|in_

### Parameters

* `string|float $value`

---

## pageMargins()

---

```php
public function pageMargins(array $margins): Sheet
```
_Page margins for a sheet or a custom sheet view in mm|cm|in_

### Parameters

* `array $margins`

---

## pageMarginTop()

---

```php
public function pageMarginTop($value): Sheet
```
_Top Page Margin in mm|cm|in_

### Parameters

* `string|float $value`

---

## pageOrientationLandscape()

---

```php
public function pageOrientationLandscape(): Sheet
```
_Set page orientation as Landscape, alias of pageLandscape()_

### Parameters

_None_

---

## pageOrientationPortrait()

---

```php
public function pageOrientationPortrait(): Sheet
```
_Set page orientation as Portrait, alias of pagePortrait()_

### Parameters

_None_

---

## pagePaperHeight()

---

```php
public function pagePaperHeight($paperHeight): Sheet
```
_Height of custom paper as a number followed by a unit identifier mm|cm|in (ex: 297mm, 11in)_

### Parameters

* `string|float|int $paperHeight`

---

## pagePaperSize()

---

```php
public function pagePaperSize(int $paperSize): Sheet
```
_Set Paper size (when paperHeight and paperWidth are specified, paperSize should be ignored)_

### Parameters

* `int $paperSize`

---

## pagePaperSizeA3()

---

```php
public function pagePaperSizeA3(): Sheet
```
_Set Paper Size to A3 (when paperHeight and paperWidth are specified, paperSize should be ignored)_

### Parameters

_None_

---

## pagePaperSizeA4()

---

```php
public function pagePaperSizeA4(): Sheet
```
_Set Paper Size to A4 (when paperHeight and paperWidth are specified, paperSize should be ignored)_

### Parameters

_None_

---

## pagePaperSizeLegal()

---

```php
public function pagePaperSizeLegal(): Sheet
```
_Set Paper Size to Legal (when paperHeight and paperWidth are specified, paperSize should be ignored)_

### Parameters

_None_

---

## pagePaperSizeLetter()

---

```php
public function pagePaperSizeLetter(): Sheet
```
_Set Paper Size to Letter (when paperHeight and paperWidth are specified, paperSize should be ignored)_

### Parameters

_None_

---

## pagePaperWidth()

---

```php
public function pagePaperWidth($paperWidth): Sheet
```
_Width of custom paper as a number followed by a unit identifier mm|cm|in (ex: 21cm, 8.5in)_

### Parameters

* `string|float|int $paperWidth`

---

## pagePortrait()

---

```php
public function pagePortrait(): Sheet
```
_Set page orientation as Portrait_

### Parameters

_None_

---

## pageScale()

---

```php
public function pageScale(int $scale): Sheet
```
_Set page scale_

### Parameters

* `int $scale`

---

## protect()

---

```php
public function protect(?string $password = null): Sheet
```
_Protect sheet_

### Parameters

* `string|null $password`

---

## setActiveCell()

---

```php
public function setActiveCell($cellAddress): Sheet
```
_Set active cell_

### Parameters

* `$cellAddress`

---

## setAutoFilter()

---

```php
public function setAutoFilter($rowOrCell, ?int $col = 1): Sheet
```
_Set auto filter_

### Parameters

* `mixed|null $rowOrCell`
* `int|null $col`

---

### Examples

```php
$sheet->setAutoFilter(2);
$sheet->setAutoFilter('B2');
$sheet->setAutoFilter('B2:C4');
```


---

## setBgColor()

---

```php
public function setBgColor(string $cellAddr, string $color): Sheet
```
_Set background color for the specific cell or range_

### Parameters

* `string $cellAddr`
* `string $color`

---

## setBottomNodesOptions()

---

```php
public function setBottomNodesOptions(string $node, array $options): Sheet
```
_Set multiple options for a bottom node_

### Parameters

* `string $node`
* `array $options`

---

## setCellStyle()

---

```php
public function setCellStyle(string $cellAddress, $style, 
                             ?bool $mergeStyles = false): Sheet
```
_Set style for the specific cell_

### Parameters

* `string $cellAddress` – Cell address
* `array|Style $style` – Style array or object
* `bool|null $mergeStyles` – True - merge style with previous style for this cell (if exists)

---

## setColAutoWidth()

---

```php
public function setColAutoWidth($col): Sheet
```
_Alias of setColWidthAuto($col)_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)

---

## setColDataStyle()

---

```php
public function setColDataStyle($colRange, $colStyle): Sheet
```
_Set styles of column cells (colors, formats, etc.)_

_Styles are applied only to non-empty cells in a column and only take effect starting with the current row_

### Parameters

* `int|string|array $colRange`
* `array|Style $colStyle`

---

### Examples

```php
$sheet->setColDataStyle('B', ['width' = 20]); // style for cells of column 'B'
$sheet->setColDataStyle(2, ['width' = 20]); // 'B' is number 2 column
$sheet->setColDataStyle('B:D', ['width' = 'auto']); // options for range of columns
$sheet->setColDataStyle(['A', 'B', 'C'], $style); // options for several columns 'A', 'B' and 'C'
```


---

## setColDataStyleArray()

---

```php
public function setColDataStyleArray(array $colStyles): Sheet
```
_Set style of column cells (colors, formats, etc.)_

_Styles are applied only to non-empty cells in a column and only take effect starting with the current row_

### Parameters

* `array $colStyles`

---

### Examples

```php
$sheet->setColDataStyleArray(['B' => $style1, 'C' => $style2]); // options for columns 'B' and 'C'
```


---

## setColFormat()

---

```php
public function setColFormat($col, $format): Sheet
```
_Set a format of single or multiple column(s)_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)
* `mixed $format`

---

## setColFormats()

---

```php
public function setColFormats(array $formats): Sheet
```
_Set formats of columns_

### Parameters

* `array $formats`

---

## setColFormula()

---

```php
public function setColFormula($col, string $formula): Sheet
```
_Set formula for single or multiple column(s)_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)
* `string $formula`

---

## setColHidden()

---

```php
public function setColHidden($col): Sheet
```
_Set a column as hidden_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)

---

## setColMinWidth()

---

```php
public function setColMinWidth($col, $width): Sheet
```
_Setting a minimal column's width_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)
* `int|float|string $width`

---

## setColMinWidths()

---

```php
public function setColMinWidths(array $widths): Sheet
```
_Setting a multiple column's minimal width_

### Parameters

* `array $widths`

---

### Examples

```php
$sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);
```


---

## setColOutlineLevel()

---

```php
public function setColOutlineLevel($col, int $outlineLevel): Sheet
```
_Set a column outline level_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)
* `int $outlineLevel`

---

## setColStyle()

---

```php
public function setColStyle($colRange, $style): Sheet
```
_Set style of single or multiple column(s)_

_Styles are applied to the entire sheet column(s) (even if it is empty)_

### Parameters

* `int|string|array $colRange` – Column number or column letter or column range (or array of these)
* `array|Style $style`

---

### Examples

```php
$sheet->setColStyle('B', $style);
$sheet->setColStyle(2, $style); // 'B' is number 2 column
$sheet->setColStyle('C:F', $style);
$sheet->setColStyle(['A', 'B', 'C'], $style);
```


---

## setColStyleArray()

---

```php
public function setColStyleArray($colStyles): Sheet
```
_Set styles of single or multiple column(s)_

_Styles are applied to the entire sheet column(s) (even if it is empty)_

### Parameters

* `array|Style $colStyles`

---

### Examples

```php
$style1 = ['width' = 20];
$style2 = (new Style())->setColor('red');
$sheet->setColStyleArray(['B' => $style1, 'C' => $style2]);
```


---

## setColVisible()

---

```php
public function setColVisible($col, bool $val): Sheet
```
_Show/hide a column_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)
* `bool $val`

---

## setColWidth()

---

```php
public function setColWidth($col, $width, ?bool $min = false): Sheet
```
_Set a width of single or multiple column(s)_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)
* `int|float|string $width`
* `bool|null $min`

---

## setColWidthAuto()

---

```php
public function setColWidthAuto($col): Sheet
```
_Set auto width of single or multiple column(s)_

### Parameters

* `int|string|array $col` – Column number or column letter (or array of these)

---

### Examples

```php
$sheet->setColWidthAuto(2);
$sheet->setColWidthAuto('B');
$sheet->setColWidthAuto(['B', 'C']);
$sheet->setColWidthAuto(['B:D']);
```


---

## setColWidths()

---

```php
public function setColWidths(array $widths, ?bool $min = false): Sheet
```
_Setting a multiple column's width_

### Parameters

* `array $widths`
* `bool|null $min`

---

### Examples

```php
$sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);
```


---

## setDefaultFont()

---

```php
public function setDefaultFont($font): Sheet
```
_Set default font settings for the sheet_

### Parameters

* `string|array $font`

---

## setDefaultFontColor()

---

```php
public function setDefaultFontColor(string $fontColor): Sheet
```
_Set default font color_

### Parameters

* `string $fontColor`

---

## setDefaultFontName()

---

```php
public function setDefaultFontName(string $fontName): Sheet
```
_Set default font name for the sheet_

### Parameters

* `string $fontName`

---

## setDefaultFontSize()

---

```php
public function setDefaultFontSize(int $fontSize): Sheet
```
_Set default font size for the sheet_

### Parameters

* `int $fontSize`

---

## setDefaultFontStyle()

---

```php
public function setDefaultFontStyle(string $fontStyle): Sheet
```
_Set default font style (bold, italic, etc.) for the sheet_

### Parameters

* `string $fontStyle`

---

## setDefaultFontStyleBold()

---

```php
public function setDefaultFontStyleBold(): Sheet
```
_Set default font style as bold for the sheet_

### Parameters

_None_

---

## setDefaultFontStyleItalic()

---

```php
public function setDefaultFontStyleItalic(): Sheet
```
_Set default font style as italic_

### Parameters

_None_

---

## setDefaultFontStyleStrikethrough()

---

```php
public function setDefaultFontStyleStrikethrough(): Sheet
```
_Set default font style as strikethrough_

### Parameters

_None_

---

## setDefaultFontStyleUnderline()

---

```php
public function setDefaultFontStyleUnderline(?bool $double = false): Sheet
```
_Set default font style as underline_

### Parameters

* `bool|null $double`

---

## setDefaultStyle()

---

```php
public function setDefaultStyle($style): Sheet
```
_Sets default style_

### Parameters

* `array|Style $style`

---

## setFormat()

---

```php
public function setFormat(string $cellAddr, string $format): Sheet
```
_Set value format for the specific cell or range_

### Parameters

* `string $cellAddr`
* `string $format`

---

## setFormula()

---

```php
public function setFormula($cellAddress, $value, $style): Sheet
```
_Set a formula to the single cell or to the cell range_

### Parameters

* `string|array $cellAddress`
* `mixed $value`
* `array|Style|null $style`

---

### Examples

```php
$sheet->setFormula('B5', '=F23');
$sheet->setFormula('B5:C7', $formula, $style);
$sheet->setFormula(['col' => 2, 'row' => 5], '=R2C3+R3C4');
$sheet->setFormula([2, 5], '=SUM(A4:A18)');
```


---

## setFreeze()

---

```php
public function setFreeze($freezeRows, $freezeColumns): Sheet
```
_Freeze rows/columns_

### Parameters

* `mixed $freezeRows`
* `mixed $freezeColumns`

---

### Examples

```php
$sheet->setFreeze(3, 3); // number rows and columns to freeze
$sheet->setFreeze('C3'); // left top cell of the free area
```


---

## setFreezeColumns()

---

```php
public function setFreezeColumns(int $freezeColumns): Sheet
```
_Freeze columns_

### Parameters

* `int $freezeColumns` – Number columns to freeze

---

## setFreezeRows()

---

```php
public function setFreezeRows(int $freezeRows): Sheet
```
_Freeze rows_

### Parameters

* `int $freezeRows` – Number rows to freeze

---

## setName()

---

```php
public function setName(string $sheetName): Sheet
```
_Set sheet name_

### Parameters

* `string $sheetName`

---

## setOuterBorder()

---

```php
public function setOuterBorder(string $range, $style): Sheet
```
_Set outer border for the specific range_

### Parameters

* `string $range`
* `string|array|Style $style`

---

## setPrintArea()

---

```php
public function setPrintArea(string $range): Sheet
```
_Set print area for the sheet_

### Parameters

* `string $range`

---

## setPrintCentered()

---

```php
public function setPrintCentered(?bool $value = true): Sheet
```
_Center the print area horizontally and vertically_

### Parameters

* `bool|null $value`

---

## setPrintGridlines()

---

```php
public function setPrintGridlines(?bool $bool = true): Sheet
```
_Show grid lines in the print area_

### Parameters

* `bool|null $bool`

---

## setPrintHorizontalCentered()

---

```php
public function setPrintHorizontalCentered(?bool $value = true): Sheet
```
_Center the print area horizontally_

### Parameters

* `bool|null $value`

---

## setPrintLeftColumns()

---

```php
public function setPrintLeftColumns(string $range): Sheet
```
_Set left columns to repeat on every printed page_

### Parameters

* `string $range`

---

## setPrintRowAndColumnHeading()

---

```php
public function setPrintRowAndColumnHeading(?bool $bool = true): Sheet
```
_Print row and column headings in the printout_

### Parameters

* `bool|null $bool`

---

## setPrintTitles()

---

```php
public function setPrintTitles(?string $rowsAtTop, 
                               ?string $colsAtLeft = null): Sheet
```
_Set rows to repeat at top and columns to repeat at left when printing_

### Parameters

* `string|null $rowsAtTop`
* `string|null $colsAtLeft`

---

## setPrintTopRows()

---

```php
public function setPrintTopRows(string $range): Sheet
```
_Set top rows to repeat on every printed page_

### Parameters

* `string $range`

---

## setPrintVerticalCentered()

---

```php
public function setPrintVerticalCentered(?bool $value = true): Sheet
```
_Center the print area vertically_

### Parameters

* `bool|null $value`

---

## setRowDataStyle()

---

```php
public function setRowDataStyle($rowRange, $style): Sheet
```
_Style are applied only to non-empty cells in a row (or row range)_

### Parameters

* `int|string|array $rowRange`
* `array|Style $style`

---

### Examples

```php
$sheet->setRowDataStyle(3, ['height' = 20]); // options for row number 3
$sheet->setRowDataStyle('2:5', ['font-color' = '#f00']); // options for range of rows
```


---

## setRowDataStyleArray()

---

```php
public function setRowDataStyleArray(array $rowStyles): Sheet
```
_Styles are applied only to non-empty cells in a rows_

### Parameters

* `array $rowStyles`

---

### Examples

```php
$sheet->setRowDataStyleArray([3 => $style1, 5 => $style2]); // styles for rows 3 and 5
```


---

## setRowHeight()

---

```php
public function setRowHeight($rowNum, $height): Sheet
```
_Height of a specific row_

### Parameters

* `$rowNum`
* `$height`

---

## setRowHeights()

---

```php
public function setRowHeights(array $heights): Sheet
```
_Multiple rows height_

### Parameters

* `array $heights`

---

## setRowHidden()

---

```php
public function setRowHidden($rowNum): Sheet
```
_Hide a specific row_

### Parameters

* `int|array $rowNum`

---

## setRowOutlineLevel()

---

```php
public function setRowOutlineLevel($rowNum, int $outlineLevel, 
                                   ?bool $collapsed = null): Sheet
```
_Set an outline level for a specific row or range of rows_

### Parameters

* `int|array|string $rowNum`
* `int $outlineLevel`
* `bool|null $collapsed`

---

### Examples

```php
setRowOutlineLevel(5, 1)
setRowOutlineLevel([5, 6, 7], 1)
setRowOutlineLevel('5:7', 1)
```


---

## setRowStyle()

---

```php
public function setRowStyle($rowRange, $style): Sheet
```
_The style is applied to the entire sheet row (even if it is empty)_

### Parameters

* `int|string|array $rowRange`
* `array|Style $style`

---

### Examples

```php
$sheet->setRowStyle(3, ['height' = 20]); // options for row number 3
$sheet->setRowStyle('2:5', ['font-color' = '#f00']); // options for range of rows
```


---

## setRowStyleArray()

---

```php
public function setRowStyleArray(array $rowStyles): Sheet
```
_Styles are applied to the entire sheet row (even if it is empty)_

### Parameters

* `array $rowStyles`

---

### Examples

```php
$sheet->setRowStyleArray([3 => $style1, 5 => $style2]); // styles for rows 3 and 5
```


---

## setRowVisible()

---

```php
public function setRowVisible($rowNum, bool $visible): Sheet
```
_Hide/show a specific row_

### Parameters

* `int|array $rowNum`
* `bool $visible`

---

## setShowGridLines()

---

```php
public function setShowGridLines(bool $flag): void
```
_Turn on/off grid lines_

### Parameters

* `bool $flag`

---

## setStateHidden()

---

```php
public function setStateHidden(): Sheet
```
_Make the sheet hidden_

### Parameters

_None_

---

## setStateVeryHidden()

---

```php
public function setStateVeryHidden(): Sheet
```
_Make the sheet very hidden_

### Parameters

_None_

---

## setStateVisible()

---

```php
public function setStateVisible(): Sheet
```
_Make the sheet visible_

### Parameters

_None_

---

## setStyle()

---

```php
public function setStyle(string $cellAddress, $style, 
                         ?bool $mergeStyles = false): Sheet
```
_Alias for 'setCellStyle()'_

### Parameters

* `string $cellAddress`
* `array|Style $style`
* `bool|null $mergeStyles`

---

## setTabColor()

---

```php
public function setTabColor(?string $color): Sheet
```
_Set color for the sheet tab_

### Parameters

* `string|null $color`

---

## setTopLeftCell()

---

```php
public function setTopLeftCell($cellAddress): Sheet
```
_Set the top left cell for writing_

### Parameters

* `string|array $cellAddress`

---

### Examples

```php
$sheet->setTopLeftCell('C3');
$sheet->writeRow([11, 22, 33]); // Will be written in cells C3, D3, E3
$sheet->setTopLeftCell('G7');
$sheet->writeRow([44, 55]); // Will be written in cells G7, H7
```


---

## setValue()

---

```php
public function setValue($cellAddress, $value, $style): Sheet
```
_Set a value to the single cell or to the cell range_

### Parameters

* `string|array $cellAddress`
* `mixed $value`
* `array|Style|null $style`

---

### Examples

```php
$sheet->setValue('B5', $value);
$sheet->setValue('B5:C7', $value, $style);
$sheet->setValue(['col' => 2, 'row' => 5], $value, $style);
$sheet->setValue([2, 5], $value);
```


---

## skipRow()

---

```php
public function skipRow(?int $rowCount = 1): Sheet
```
_Skip rows_

### Parameters

* `int|null $rowCount`

---

## unprotect()

---

```php
public function unprotect(): Sheet
```
_Unprotect sheet_

### Parameters

_None_

---

## withLastCell()

---

```php
public function withLastCell(): Sheet
```
_Select the last written cell for applying_

### Parameters

_None_

---

## withLastRow()

---

```php
public function withLastRow(): Sheet
```
_Select the last written row for applying_

### Parameters

_None_

---

## withRange()

---

```php
public function withRange($range): Sheet
```
_Select a custom range for applying_

### Parameters

* `array|string $range`

---

## writeArray()

---

```php
public function writeArray(array $rowArray = [], $rowStyle): Sheet
```
_Write values from a two-dimensional array (alias of writeRows)_

### Parameters

* `array $rowArray` – Array of rows
* `array|Style|null $rowStyle` – Style applied to each row

---

## writeArrayTo()

---

```php
public function writeArrayTo($topLeftCell, array $data): Sheet
```
_Write 2d array from the specified cell_

### Parameters

* `$topLeftCell`
* `array $data`

---

## writeCell()

---

```php
public function writeCell($value, ?array $styles = null): Sheet
```
_Write value to the current cell and move a pointer to the next cell in the row_

### Parameters

* `mixed $value`
* `array|null $styles`

---

## writeCells()

---

```php
public function writeCells(array $values, ?array $cellStyles = null): Sheet
```
_Write several values into cells of one row_

### Parameters

* `array $values`
* `array|null $cellStyles`

---

## writeHeader()

---

```php
public function writeHeader(array $header, ?array $rowStyle = null, 
                            ?array $colStyles = []): Sheet
```
_Write a header row with optional styles and formats for columns_

### Parameters

* `array $header`
* `array|null $rowStyle`
* `array|null $colStyles`

---

### Examples

```php
$sheet->writeHeader(['title1', 'title2', 'title3']); // texts for cells of header
$sheet->writeHeader(['title1' => '@text', 'title2' => 'YYYY-MM-DD', 'title3' => ['format' => ..., 'font' => ...]]); // texts and formats of columns
$sheet->writeHeader($cellValues, $rowStyle, $colStyles); // texts and formats of columns and options of row
```


---

## writeRow()

---

```php
public function writeRow(array $rowValues = [], $rowStyle, 
                         ?array $cellStyles = null): Sheet
```
_Write values to the current row_

### Parameters

* `array $rowValues` – Values of cells
* `array|Style|null $rowStyle` – Style applied to the entire row
* `array|null $cellStyles` – Styles of specified cells in the row

---

## writeRows()

---

```php
public function writeRows(array $rowArray = [], 
                          ?array $rowStyle = null): Sheet
```
_Write several rows from a two-dimensional array_

### Parameters

* `array $rowArray` – Array of rows
* `array|null $rowStyle` – Style applied to each row

---

## writeTo()

---

```php
public function writeTo($cellAddress, $value, $style, 
                        ?int $mergeFlag = 0): Sheet
```
_Write value to the specified cell and move a pointer to the next cell in the row_

### Parameters

* `string|array $cellAddress`
* `mixed $value`
* `array|Style|null $style`
* `int|null $mergeFlag`

---

### Examples

```php
$sheet->writeTo('B5', $value); // write to single cell
$sheet->writeTo(['col' => 2, 'row' => 5], $value); // address as an array
$sheet->writeTo([2, 5], $value); // address as an array
$sheet->writeTo('B5:C7', $value); // write a value to merged cells
$sheet->writeTo('B5:C7', $value, $styles, Sheet:MERGE_NO_CHECK); // don't check for intersection of merged cells
```


---

