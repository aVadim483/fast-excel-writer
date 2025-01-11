# Class \avadim\FastExcelWriter\Sheet

---

* [__construct()](#__construct) -- Sheet constructor
* [setActiveCell()](#setactivecell) -- Set active cell
* [addCellStyle()](#addcellstyle) -- Add additional styles to a cell
* [addChart()](#addchart) -- Add chart object to the specified range of cells
* [addConditionalFormatting()](#addconditionalformatting) -- Add conditional formatting object to the specified range of cells
* [addDataValidation()](#adddatavalidation) -- Add data validation object to the specified range of cells
* [addImage()](#addimage) -- Add image to the sheet from local file, URL or image string in base64
* [addNamedRange()](#addnamedrange) -- Define named range
* [addNote()](#addnote) -- Add note to the sheet
* [addStyle()](#addstyle) -- Alias for 'addCellStyle()'
* [allowAutoFilter()](#allowautofilter) -- AutoFilters should be allowed to operate when the sheet is protected
* [allowDeleteColumns()](#allowdeletecolumns) -- Deleting columns should be allowed when the sheet is protected
* [allowDeleteRows()](#allowdeleterows) -- Deleting rows should be allowed when the sheet is protected
* [allowEditObjects()](#alloweditobjects) -- Objects are allowed to be edited when the sheet is protected
* [allowEditScenarios()](#alloweditscenarios) -- Scenarios are allowed to be edited when the sheet is protected
* [allowFormatCells()](#allowformatcells) -- Formatting cells should be allowed when the sheet is protected
* [allowFormatColumns()](#allowformatcolumns) -- Formatting columns should be allowed when the sheet is protected
* [allowFormatRows()](#allowformatrows) -- Formatting rows should be allowed when the sheet is protected
* [allowInsertColumns()](#allowinsertcolumns) -- Inserting columns should be allowed when the sheet is protected
* [allowInsertHyperlinks()](#allowinserthyperlinks) -- Inserting hyperlinks should be allowed when the sheet is protected
* [allowInsertRows()](#allowinsertrows) -- Inserting rows should be allowed when the sheet is protected
* [allowPivotTables()](#allowpivottables) -- PivotTables should be allowed to operate when the sheet is protected
* [allowSelectCells()](#allowselectcells) -- Selection of any cells should be allowed when the sheet is protected
* [allowSelectLockedCells()](#allowselectlockedcells) -- Selection of locked cells should be allowed when the sheet is protected
* [allowSelectUnlockedCells()](#allowselectunlockedcells) -- Selection of unlocked cells should be allowed when the sheet is protected
* [allowSort()](#allowsort) -- Sorting should be allowed when the sheet is protected
* [applyAlignLeft()](#applyalignleft) -- Apply left alignment to content
* [applyAlignRight()](#applyalignright) -- Apply right alignment to content
* [applyBgColor()](#applybgcolor) -- Alias of 'applyFillColor()'
* [applyBorder()](#applyborder) -- Sets all borders style
* [applyBorderBottom()](#applyborderbottom)
* [applyBorderLeft()](#applyborderleft)
* [applyBorderRight()](#applyborderright)
* [applyBorderTop()](#applybordertop)
* [applyColor()](#applycolor) -- Alias of 'setFontColor()'
* [applyConditionalFormatting()](#applyconditionalformatting)
* [applyDataValidation()](#applydatavalidation)
* [applyFillColor()](#applyfillcolor) -- Fill background color
* [applyFillGradient()](#applyfillgradient) -- Fill background by gradient
* [applyFont()](#applyfont)
* [applyFontColor()](#applyfontcolor)
* [applyFontName()](#applyfontname)
* [applyFontSize()](#applyfontsize)
* [applyFontStyle()](#applyfontstyle)
* [applyFontStyleBold()](#applyfontstylebold)
* [applyFontStyleItalic()](#applyfontstyleitalic)
* [applyFontStyleStrikethrough()](#applyfontstylestrikethrough)
* [applyFontStyleUnderline()](#applyfontstyleunderline)
* [applyFormat()](#applyformat)
* [applyHide()](#applyhide)
* [applyIndentDistributed()](#applyindentdistributed)
* [applyIndentLeft()](#applyindentleft)
* [applyIndentRight()](#applyindentright)
* [applyInnerBorder()](#applyinnerborder)
* [applyNamedRange()](#applynamedrange)
* [applyOuterBorder()](#applyouterborder)
* [applyRowHeight()](#applyrowheight) -- Sets height to the current row
* [applyRowOutlineLevel()](#applyrowoutlinelevel)
* [applyStyle()](#applystyle)
* [applyTextAlign()](#applytextalign)
* [applyTextCenter()](#applytextcenter)
* [applyTextColor()](#applytextcolor)
* [applyTextRotation()](#applytextrotation)
* [applyTextWrap()](#applytextwrap)
* [applyUnlock()](#applyunlock)
* [applyVerticalAlign()](#applyverticalalign)
* [setAutofilter()](#setautofilter) -- Set auto filter
* [beginArea()](#beginarea) -- Begin a new area
* [beginOutlineLevel()](#beginoutlinelevel)
* [setBgColor()](#setbgcolor)
* [getBottomNodesOptions()](#getbottomnodesoptions)
* [setBottomNodesOptions()](#setbottomnodesoptions)
* [cell()](#cell) -- Select a single cell or cell range in the current row
* [setCellStyle()](#setcellstyle) -- Set style for the specific cell
* [getCharts()](#getcharts)
* [clearAreas()](#clearareas)
* [getColAttributes()](#getcolattributes)
* [setColAutoWidth()](#setcolautowidth)
* [setColDataStyle()](#setcoldatastyle) -- Set style of column cells (colors, formats, etc.)
* [setColDataStyleArray()](#setcoldatastylearray) -- Set style of column cells (colors, formats, etc.)
* [setColFormat()](#setcolformat) -- Set format of single or multiple column(s)
* [setColFormats()](#setcolformats) -- Set formats of columns
* [setColFormula()](#setcolformula) -- Set formula for single or multiple column(s)
* [setColHidden()](#setcolhidden)
* [setColMinWidth()](#setcolminwidth) -- Setting a minimal column's width
* [setColMinWidths()](#setcolminwidths) -- Setting a multiple column's minimal width
* [setColOptions()](#setcoloptions) -- Use 'setColDataStyle()' or 'setColDataStyleArray()' instead
* [setColOutlineLevel()](#setcoloutlinelevel)
* [setColStyle()](#setcolstyle) -- Set style of single or multiple column(s)
* [setColStyleArray()](#setcolstylearray) -- Set style of single or multiple column(s)
* [setColStyles()](#setcolstyles)
* [setColVisible()](#setcolvisible) -- Show/hide a column
* [setColWidth()](#setcolwidth) -- Set width of single or multiple column(s)
* [setColWidthAuto()](#setcolwidthauto) -- Set width of single or multiple column(s)
* [setColWidths()](#setcolwidths) -- Setting a multiple column's width
* [getConditionalFormatting()](#getconditionalformatting)
* [getCurrentCell()](#getcurrentcell) -- Returns address of the current cell
* [getCurrentCol()](#getcurrentcol) -- Returns current column letter
* [getCurrentColId()](#getcurrentcolid)
* [getCurrentRow()](#getcurrentrow) -- Returns current row number
* [getCurrentRowId()](#getcurrentrowid)
* [getDataValidations()](#getdatavalidations)
* [setDefaultFont()](#setdefaultfont)
* [setDefaultFontColor()](#setdefaultfontcolor)
* [setDefaultFontName()](#setdefaultfontname)
* [setDefaultFontSize()](#setdefaultfontsize)
* [setDefaultFontStyle()](#setdefaultfontstyle)
* [setDefaultFontStyleBold()](#setdefaultfontstylebold)
* [setDefaultFontStyleItalic()](#setdefaultfontstyleitalic)
* [setDefaultFontStyleStrikethrough()](#setdefaultfontstylestrikethrough)
* [setDefaultFontStyleUnderline()](#setdefaultfontstyleunderline)
* [getDefaultStyle()](#getdefaultstyle) -- Returns default style
* [setDefaultStyle()](#setdefaultstyle) -- Sets default style
* [endAreas()](#endareas)
* [endOutlineLevel()](#endoutlinelevel)
* [setFormat()](#setformat)
* [setFormula()](#setformula) -- Set a formula to the single cell or to the cell range
* [setFreeze()](#setfreeze) -- Freeze rows/columns
* [setFreezeColumns()](#setfreezecolumns) -- Freeze columns
* [setFreezeRows()](#setfreezerows) -- Freeze rows
* [getImages()](#getimages)
* [getLastCell()](#getlastcell)
* [getLastRange()](#getlastrange)
* [makeArea()](#makearea) -- Make area for writing
* [mergeCells()](#mergecells) -- Merge cells
* [getMergedCells()](#getmergedcells) -- Returns merged cells
* [mergeRelCells()](#mergerelcells) -- Merge relative cells
* [getName()](#getname) -- Get sheet name
* [isName()](#isname) -- Case-insensitive name checking
* [setName()](#setname) -- Set sheet name
* [getNamedRanges()](#getnamedranges) -- Returns named ranges with full addresses
* [nextCell()](#nextcell)
* [nextRow()](#nextrow) -- Move to the next row
* [getNotes()](#getnotes)
* [setOuterBorder()](#setouterborder)
* [getOutlineLevel()](#getoutlinelevel)
* [getPageFit()](#getpagefit)
* [pageFitToHeight()](#pagefittoheight)
* [getPageFitToHeight()](#getpagefittoheight)
* [pageFitToWidth()](#pagefittowidth)
* [getPageFitToWidth()](#getpagefittowidth)
* [pageLandscape()](#pagelandscape) -- Set page orientation as Landscape
* [pageMarginBottom()](#pagemarginbottom) -- Bottom Page Margin in mm|cm|in
* [pageMarginFooter()](#pagemarginfooter) -- Footer Page Margin in mm|cm|in
* [pageMarginHeader()](#pagemarginheader) -- Header Page Margin in mm|cm|in
* [pageMarginLeft()](#pagemarginleft) -- Left Page Margin in mm|cm|in
* [pageMarginRight()](#pagemarginright) -- Right page margin in mm|cm|in
* [pageMargins()](#pagemargins) -- Page margins for a sheet or a custom sheet view in mm|cm|in
* [getPageMargins()](#getpagemargins)
* [setPageMargins()](#setpagemargins)
* [pageMarginTop()](#pagemargintop) -- Top Page Margin in mm|cm|in
* [setPageOptions()](#setpageoptions)
* [getPageOrientation()](#getpageorientation)
* [pageOrientationLandscape()](#pageorientationlandscape) -- Set page orientation as Landscape, alias of pageLandscape()
* [pageOrientationPortrait()](#pageorientationportrait) -- Set page orientation as Portrait, alias of pagePortrait()
* [pagePaperHeight()](#pagepaperheight) -- Height of custom paper as a number followed by a unit identifier mm|cm|in (ex: 297mm, 11in)
* [pagePaperSize()](#pagepapersize) -- Set Paper size (when paperHeight and paperWidth are specified, paperSize should be ignored)
* [pagePaperWidth()](#pagepaperwidth) -- Width of custom paper as a number followed by a unit identifier mm|cm|in (ex: 21cm, 8.5in)
* [pagePortrait()](#pageportrait) -- Set page orientation as Portrait
* [pageScale()](#pagescale)
* [getPageSetup()](#getpagesetup)
* [setPageSetup()](#setpagesetup)
* [setPrintArea()](#setprintarea)
* [setPrintCentered()](#setprintcentered) -- Center the print area horizontally and vertically
* [setPrintGridlines()](#setprintgridlines) -- Show grid line in the print area
* [setPrintHorizontalCentered()](#setprinthorizontalcentered) -- Center the print area horizontally
* [setPrintLeftColumns()](#setprintleftcolumns)
* [setPrintTitles()](#setprinttitles)
* [setPrintTopRows()](#setprinttoprows)
* [setPrintVerticalCentered()](#setprintverticalcentered) -- Center the print area vertically
* [protect()](#protect) -- Protect sheet
* [isRightToLeft()](#isrighttoleft)
* [setRowDataStyle()](#setrowdatastyle) -- Style are applied only to non-empty cells in a row (or row range)
* [setRowDataStyleArray()](#setrowdatastylearray) -- Styles are applied only to non-empty cells in a rows
* [setRowHeight()](#setrowheight) -- Height of a specific row
* [setRowHeights()](#setrowheights) -- Multiple rows height
* [setRowHidden()](#setrowhidden) -- Hide a specific row
* [setRowOutlineLevel()](#setrowoutlinelevel)
* [setRowStyle()](#setrowstyle) -- Style are applied to the entire sheet row (even if it is empty)
* [setRowStyleArray()](#setrowstylearray) -- Styles are applied to the entire sheet row (even if it is empty)
* [setRowStyles()](#setrowstyles)
* [setRowVisible()](#setrowvisible) -- Hide/show a specific row
* [skipRow()](#skiprow) -- Skip rows
* [setStyle()](#setstyle) -- Alias for 'setCellStyle()'
* [setTabColor()](#settabcolor) -- Set color for the sheet tab
* [setTopLeftCell()](#settopleftcell) -- Set top left cell for writing
* [unprotect()](#unprotect) -- Unprotect sheet
* [setValue()](#setvalue) -- Set a value to the single cell or to the cell range
* [withLastCell()](#withlastcell) -- Select last written cell for applying
* [withLastRow()](#withlastrow) -- Select last written row for applying
* [withRange()](#withrange) -- Select custom range for applying
* [writeAreas()](#writeareas)
* [writeArray()](#writearray) -- Write values from two-dimensional array
* [writeArrayTo()](#writearrayto) -- Write 2d array form the specified cell
* [writeCell()](#writecell) -- Write value to the current cell and move pointer to the next cell in the row
* [writeCells()](#writecells) -- Write several values into cells of one row
* [writeHeader()](#writeheader)
* [writeRow()](#writerow) -- Write values to the current row
* [writeTo()](#writeto) -- Write value to the specified cell and move pointer to the next cell in the row

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

## setActiveCell()

---

```php
public function setActiveCell($cellAddress): Sheet
```
_Set active cell_

### Parameters

* `$cellAddress`

---

## addCellStyle()

---

```php
public function addCellStyle(string $cellAddr, array $style): Sheet
```
_Add additional styles to a cell_

### Parameters

* `string $cellAddr`
* `array $style`

---

## addChart()

---

```php
public function addChart(string $range, 
                         avadim\FastExcelWriter\Charts\Chart $chart): Sheet
```
_Add chart object to the specified range of cells_

### Parameters

* `string $range` -- Set the position where the chart should appear in the worksheet
* `Chart $chart` -- Chart object

---

## addConditionalFormatting()

---

```php
public function addConditionalFormatting(string $range, $conditionals): Sheet
```
_Add conditional formatting object to the specified range of cells_

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
_Add data validation object to the specified range of cells_

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
_Add image to the sheet from local file, URL or image string in base64_

### Parameters

* `string $cell` -- Cell address
* `string $imageFile` -- URL, local path or image string in base64
* `array|null $imageStyle` -- \['width' => ..., 'height' => ..., 'hyperlink' => ...]

---

### Examples

```php
$sheet->addImage('A1', 'path/to/file');
$sheet->addImage('A1', 'path/to/file', ['width => 100]);
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
_Add note to the sheet_

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
public function addStyle(string $cellAddr, array $style): Sheet
```
_Alias for 'addCellStyle()'_

### Parameters

* `string $cellAddr`
* `array $style`

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


### Parameters

* `Conditional|Conditional[] $conditionals`

---

## applyDataValidation()

---

```php
public function applyDataValidation(avadim\FastExcelWriter\DataValidation\DataValidation $validation): Sheet
```


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


### Parameters

* `string $fontColor`

---

## applyFontName()

---

```php
public function applyFontName(string $fontName): Sheet
```


### Parameters

* `string $fontName`

---

## applyFontSize()

---

```php
public function applyFontSize(float $fontSize): Sheet
```


### Parameters

* `float $fontSize`

---

## applyFontStyle()

---

```php
public function applyFontStyle(string $fontStyle): Sheet
```


### Parameters

* `string $fontStyle`

---

## applyFontStyleBold()

---

```php
public function applyFontStyleBold(): Sheet
```


### Parameters

_None_

---

## applyFontStyleItalic()

---

```php
public function applyFontStyleItalic(): Sheet
```


### Parameters

_None_

---

## applyFontStyleStrikethrough()

---

```php
public function applyFontStyleStrikethrough(): Sheet
```


### Parameters

_None_

---

## applyFontStyleUnderline()

---

```php
public function applyFontStyleUnderline(?bool $double = false): Sheet
```


### Parameters

* `bool|null $double`

---

## applyFormat()

---

```php
public function applyFormat($format): Sheet
```


### Parameters

* `string|array $format`

---

## applyHide()

---

```php
public function applyHide(?bool $hide = true): Sheet
```


### Parameters

* `$hide`

---

## applyIndentDistributed()

---

```php
public function applyIndentDistributed(int $indent): Sheet
```


### Parameters

* `int $indent`

---

## applyIndentLeft()

---

```php
public function applyIndentLeft(int $indent): Sheet
```


### Parameters

* `int $indent`

---

## applyIndentRight()

---

```php
public function applyIndentRight(int $indent): Sheet
```


### Parameters

* `int $indent`

---

## applyInnerBorder()

---

```php
public function applyInnerBorder(string $style, 
                                 ?string $color = '#000000'): Sheet
```


### Parameters

* `string $style`
* `string|null $color`

---

## applyNamedRange()

---

```php
public function applyNamedRange(string $name): Sheet
```


### Parameters

* `string $name`

---

## applyOuterBorder()

---

```php
public function applyOuterBorder(string $style, 
                                 ?string $color = '#000000'): Sheet
```


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


### Parameters

* `int $outlineLevel`

---

## applyStyle()

---

```php
public function applyStyle(array $style): Sheet
```


### Parameters

* `array $style`

---

## applyTextAlign()

---

```php
public function applyTextAlign(string $textAlign, 
                               ?string $verticalAlign = null): Sheet
```


### Parameters

* `string $textAlign`
* `string|null $verticalAlign`

---

## applyTextCenter()

---

```php
public function applyTextCenter(): Sheet
```


### Parameters

_None_

---

## applyTextColor()

---

```php
public function applyTextColor(string $color): Sheet
```


### Parameters

* `string $color`

---

## applyTextRotation()

---

```php
public function applyTextRotation(int $degrees): Sheet
```


### Parameters

* `int $degrees`

---

## applyTextWrap()

---

```php
public function applyTextWrap(?bool $textWrap = true): Sheet
```


### Parameters

* `bool|null $textWrap`

---

## applyUnlock()

---

```php
public function applyUnlock(?bool $unlock = true): Sheet
```


### Parameters

* `bool $unlock`

---

## applyVerticalAlign()

---

```php
public function applyVerticalAlign(string $verticalAlign): Sheet
```


### Parameters

* `string $verticalAlign`

---

## setAutofilter()

---

```php
public function setAutofilter(?int $row = 1, ?int $col = 1): Sheet
```
_Set auto filter_

### Parameters

* `int|null $row`
* `int|null $col`

---

## beginArea()

---

```php
public function beginArea(?string $cellAddress = null): Area
```
_Begin a new area_

### Parameters

* `string|null $cellAddress` -- Upper left cell of area

---

## beginOutlineLevel()

---

```php
public function beginOutlineLevel(?bool $collapsed = false): Sheet
```


### Parameters

* `$collapsed`

---

## setBgColor()

---

```php
public function setBgColor(string $cellAddr, string $color): Sheet
```


### Parameters

* `string $cellAddr`
* `string $color`

---

## getBottomNodesOptions()

---

```php
public function getBottomNodesOptions(): array
```


### Parameters

_None_

---

## setBottomNodesOptions()

---

```php
public function setBottomNodesOptions(string $node, array $options): Sheet
```


### Parameters

* `string $node`
* `array $options`

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

## setCellStyle()

---

```php
public function setCellStyle(string $cellAddress, $style, 
                             ?bool $mergeStyles = false): Sheet
```
_Set style for the specific cell_

### Parameters

* `string $cellAddress` -- Cell address
* `mixed $style` -- Style array or object
* `bool|null $mergeStyles` -- True - merge style with previous style for this cell (if exists)

---

## getCharts()

---

```php
public function getCharts(): array
```


### Parameters

_None_

---

## clearAreas()

---

```php
public function clearAreas(): Sheet
```


### Parameters

_None_

---

## getColAttributes()

---

```php
public function getColAttributes(): array
```


### Parameters

_None_

---

## setColAutoWidth()

---

```php
public function setColAutoWidth($col): Sheet
```


### Parameters

* `int|string|array $col` -- Column number or column letter (or array of these)

---

## setColDataStyle()

---

```php
public function setColDataStyle($colRange, array $colStyle): Sheet
```
_Set style of column cells (colors, formats, etc.)_

_Styles are applied only to non-empty cells in a column and only take effect starting with the current row_

### Parameters

* `int|string|array $colRange`
* `array $colStyle`

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
_Set format of single or multiple column(s)_

### Parameters

* `int|string|array $col` -- Column number or column letter (or array of these)
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

* `int|string|array $col` -- Column number or column letter (or array of these)
* `string $formula`

---

## setColHidden()

---

```php
public function setColHidden($col): Sheet
```


### Parameters

* `int|string|array $col` -- Column number or column letter (or array of these)

---

## setColMinWidth()

---

```php
public function setColMinWidth($col, $width): Sheet
```
_Setting a minimal column's width_

### Parameters

* `int|string|array $col` -- Column number or column letter (or array of these)
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

## setColOptions()

---

```php
public function setColOptions($arg1, ?array $arg2 = null): Sheet
```
_Use 'setColDataStyle()' or 'setColDataStyleArray()' instead_

### Parameters

* `$arg1`
* `$arg2`

---

## setColOutlineLevel()

---

```php
public function setColOutlineLevel($col, int $outlineLevel): Sheet
```


### Parameters

* `int|string|array $col` -- Column number or column letter (or array of these)
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

* `int|string|array $colRange` -- Column number or column letter (or array of these)
* `mixed $style`

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
public function setColStyleArray(array $colStyles): Sheet
```
_Set style of single or multiple column(s)_

_Styles are applied to the entire sheet column(s) (even if it is empty)_

### Parameters

* `array $colStyles`

---

### Examples

```php
$sheet->setColStyleArray(['B' => ['width' = 20], 'C' => ['font-color' = '#f00']]);
```


---

## setColStyles()

---

```php
public function setColStyles($arg1, ?array $arg2 = null): Sheet
```


### Parameters

* `$arg1`
* `array|null $arg2`

---

## setColVisible()

---

```php
public function setColVisible($col, bool $val): Sheet
```
_Show/hide a column_

### Parameters

* `int|string|array $col` -- Column number or column letter (or array of these)
* `bool $val`

---

## setColWidth()

---

```php
public function setColWidth($col, $width, ?bool $min = false): Sheet
```
_Set width of single or multiple column(s)_

### Parameters

* `int|string|array $col` -- Column number or column letter (or array of these)
* `int|float|string $width`
* `bool|null $min`

---

## setColWidthAuto()

---

```php
public function setColWidthAuto($col): Sheet
```
_Set width of single or multiple column(s)_

### Parameters

* `int|string|array $col` -- Column number or column letter (or array of these)

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

## getConditionalFormatting()

---

```php
public function getConditionalFormatting(): array
```


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


### Parameters

_None_

---

## getDataValidations()

---

```php
public function getDataValidations(): array
```


### Parameters

_None_

---

## setDefaultFont()

---

```php
public function setDefaultFont($font): Sheet
```


### Parameters

* `string|array $font`

---

## setDefaultFontColor()

---

```php
public function setDefaultFontColor(string $fontColor): Sheet
```


### Parameters

* `string $fontColor`

---

## setDefaultFontName()

---

```php
public function setDefaultFontName(string $fontName): Sheet
```


### Parameters

* `string $fontName`

---

## setDefaultFontSize()

---

```php
public function setDefaultFontSize(int $fontSize): Sheet
```


### Parameters

* `int $fontSize`

---

## setDefaultFontStyle()

---

```php
public function setDefaultFontStyle(string $fontStyle): Sheet
```


### Parameters

* `string $fontStyle`

---

## setDefaultFontStyleBold()

---

```php
public function setDefaultFontStyleBold(): Sheet
```


### Parameters

_None_

---

## setDefaultFontStyleItalic()

---

```php
public function setDefaultFontStyleItalic(): Sheet
```


### Parameters

_None_

---

## setDefaultFontStyleStrikethrough()

---

```php
public function setDefaultFontStyleStrikethrough(): Sheet
```


### Parameters

_None_

---

## setDefaultFontStyleUnderline()

---

```php
public function setDefaultFontStyleUnderline(?bool $double = false): Sheet
```


### Parameters

* `bool|null $double`

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

## setDefaultStyle()

---

```php
public function setDefaultStyle(array $style): Sheet
```
_Sets default style_

### Parameters

* `array $style`

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


### Parameters

_None_

---

## setFormat()

---

```php
public function setFormat(string $cellAddr, string $format): Sheet
```


### Parameters

* `string $cellAddr`
* `string $format`

---

## setFormula()

---

```php
public function setFormula($cellAddress, $value, ?array $styles = null): Sheet
```
_Set a formula to the single cell or to the cell range_

### Parameters

* `string|array $cellAddress`
* `mixed $value`
* `array|null $styles`

---

### Examples

```php
$sheet->setFormula('B5', '=F23');
$sheet->setFormula('B5:C7', $formula, $styles);
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
$sheet->setFreeze('C3'); // left top cell of free area
```


---

## setFreezeColumns()

---

```php
public function setFreezeColumns(int $freezeColumns): Sheet
```
_Freeze columns_

### Parameters

* `int $freezeColumns` -- Number columns to freeze

---

## setFreezeRows()

---

```php
public function setFreezeRows(int $freezeRows): Sheet
```
_Freeze rows_

### Parameters

* `int $freezeRows` -- Number rows to freeze

---

## getImages()

---

```php
public function getImages(): array
```


### Parameters

_None_

---

## getLastCell()

---

```php
public function getLastCell(?bool $absolute = false): string
```


### Parameters

* `bool|null $absolute`

---

## getLastRange()

---

```php
public function getLastRange(?bool $absolute = false): string
```


### Parameters

* `bool|null $absolute`

---

## makeArea()

---

```php
public function makeArea(string $range): Area
```
_Make area for writing_

### Parameters

* `string $range` -- A1:Z9 or R1C1:R9C28

---

## mergeCells()

---

```php
public function mergeCells($rangeSet, ?int $actionMode = 0): Sheet
```
_Merge cells_

### Parameters

* `array|string|int $rangeSet`
* `int|null $actionMode` -- Action in case of intersection: 0 - exception; 1 - replace; 2 - keep; -1 - skip intersection check

---

### Examples

```php
$sheet->mergeCells('A1:C3');
$sheet->mergeCells(['A1:B2', 'C1:D2']);
```


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

## getName()

---

```php
public function getName(): string
```
_Get sheet name_

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

## setName()

---

```php
public function setName(string $sheetName): Sheet
```
_Set sheet name_

### Parameters

* `string $sheetName`

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

## nextCell()

---

```php
public function nextCell(): Sheet
```


### Parameters

_None_

---

## nextRow()

---

```php
public function nextRow(?array $style = []): Sheet
```
_Move to the next row_

### Parameters

* `array|null $style`

---

## getNotes()

---

```php
public function getNotes(): array
```


### Parameters

_None_

---

## setOuterBorder()

---

```php
public function setOuterBorder(string $range, $style): Sheet
```


### Parameters

* `string $range`
* `string|array $style`

---

## getOutlineLevel()

---

```php
public function getOutlineLevel(): int
```


### Parameters

_None_

---

## getPageFit()

---

```php
public function getPageFit(): bool
```


### Parameters

_None_

---

## pageFitToHeight()

---

```php
public function pageFitToHeight($numPage): Sheet
```


### Parameters

* `int|string|null $numPage`

---

## getPageFitToHeight()

---

```php
public function getPageFitToHeight(): int
```


### Parameters

_None_

---

## pageFitToWidth()

---

```php
public function pageFitToWidth($numPage): Sheet
```


### Parameters

* `int|string|null $numPage`

---

## getPageFitToWidth()

---

```php
public function getPageFitToWidth(): int
```


### Parameters

_None_

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

## getPageMargins()

---

```php
public function getPageMargins(): array
```


### Parameters

_None_

---

## setPageMargins()

---

```php
public function setPageMargins(array $margins): Sheet
```


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

## setPageOptions()

---

```php
public function setPageOptions(string $option, $value): Sheet
```


### Parameters

* `string $option`
* `mixed $value`

---

## getPageOrientation()

---

```php
public function getPageOrientation(): string
```


### Parameters

_None_

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


### Parameters

* `int $scale`

---

## getPageSetup()

---

```php
public function getPageSetup(): array
```


### Parameters

_None_

---

## setPageSetup()

---

```php
public function setPageSetup(array $options): Sheet
```


### Parameters

* `array $options`

---

## setPrintArea()

---

```php
public function setPrintArea(string $range): Sheet
```


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
_Show grid line in the print area_

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


### Parameters

* `string $range`

---

## setPrintTitles()

---

```php
public function setPrintTitles(?string $rowsAtTop, 
                               ?string $colsAtLeft = null): Sheet
```


### Parameters

* `string|null $rowsAtTop`
* `string|null $colsAtLeft`

---

## setPrintTopRows()

---

```php
public function setPrintTopRows(string $range): Sheet
```


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

## protect()

---

```php
public function protect(?string $password = null): Sheet
```
_Protect sheet_

### Parameters

* `string|null $password`

---

## isRightToLeft()

---

```php
public function isRightToLeft(): bool
```


### Parameters

_None_

---

## setRowDataStyle()

---

```php
public function setRowDataStyle($rowRange, array $style): Sheet
```
_Style are applied only to non-empty cells in a row (or row range)_

### Parameters

* `int|string|array $rowRange`
* `array $style`

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
public function setRowStyle($rowRange, array $style): Sheet
```
_Style are applied to the entire sheet row (even if it is empty)_

### Parameters

* `int|string|array $rowRange`
* `array $style`

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

## setRowStyles()

---

```php
public function setRowStyles($arg1, ?array $arg2 = null): Sheet
```


### Parameters

* `$arg1`
* `array|null $arg2`

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

## skipRow()

---

```php
public function skipRow(?int $rowCount = 1): Sheet
```
_Skip rows_

### Parameters

* `int|null $rowCount`

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
* `mixed $style`
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
_Set top left cell for writing_

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

## unprotect()

---

```php
public function unprotect(): Sheet
```
_Unprotect sheet_

### Parameters

_None_

---

## setValue()

---

```php
public function setValue($cellAddress, $value, ?array $styles = null): Sheet
```
_Set a value to the single cell or to the cell range_

### Parameters

* `string|array $cellAddress`
* `mixed $value`
* `array|null $styles`

---

### Examples

```php
$sheet->setValue('B5', $value);
$sheet->setValue('B5:C7', $value, $styles);
$sheet->setValue(['col' => 2, 'row' => 5], $value, $styles);
$sheet->setValue([2, 5], $value);
```


---

## withLastCell()

---

```php
public function withLastCell(): Sheet
```
_Select last written cell for applying_

### Parameters

_None_

---

## withLastRow()

---

```php
public function withLastRow(): Sheet
```
_Select last written row for applying_

### Parameters

_None_

---

## withRange()

---

```php
public function withRange($range): Sheet
```
_Select custom range for applying_

### Parameters

* `array|string $range`

---

## writeAreas()

---

```php
public function writeAreas(): Sheet
```


### Parameters

_None_

---

## writeArray()

---

```php
public function writeArray(array $rowArray = [], 
                           ?array $rowStyle = null): Sheet
```
_Write values from two-dimensional array_

### Parameters

* `array $rowArray` -- Array of rows
* `array|null $rowStyle` -- Style applied to each row

---

## writeArrayTo()

---

```php
public function writeArrayTo($topLeftCell, array $data): Sheet
```
_Write 2d array form the specified cell_

### Parameters

* `$topLeftCell`
* `array $data`

---

## writeCell()

---

```php
public function writeCell($value, ?array $styles = null): Sheet
```
_Write value to the current cell and move pointer to the next cell in the row_

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
public function writeRow(array $rowValues = [], ?array $rowStyle = null, 
                         ?array $cellStyles = null): Sheet
```
_Write values to the current row_

### Parameters

* `array $rowValues` -- Values of cells
* `array|null $rowStyle` -- Style applied to the entire row
* `array|null $cellStyles` -- Styles of specified cells in the row

---

## writeTo()

---

```php
public function writeTo($cellAddress, $value, ?array $styles = []): Sheet
```
_Write value to the specified cell and move pointer to the next cell in the row_

### Parameters

* `string|array $cellAddress`
* `mixed $value`
* `array|null $styles`

---

### Examples

```php
$sheet->writeTo('B5', $value); // write to single cell
$sheet->writeTo('B5:C7', $value); // write a value to merged cells
$sheet->writeTo(['col' => 2, 'row' => 5], $value); // address as array
$sheet->writeTo([2, 5], $value); // address as array
```


---

