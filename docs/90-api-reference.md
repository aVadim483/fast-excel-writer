* [Class \avadim\FastExcelWriter\Excel](#class-avadimfastexcelwriterexcel)
  * [__construct()](#__construct) Excel constructor
  * [cellAddress()](#celladdress) Create cell address by row and col numbers
  * [colIndex()](#colindex) Convert letter to index (ZERO based)
  * [colIndexRange()](#colindexrange) Convert letter range to array of numbers (ZERO based)
  * [colKeysToIndexes()](#colkeystoindexes)
  * [colKeysToLetters()](#colkeystoletters)
  * [colKeysToNumbers()](#colkeystonumbers)
  * [colLetter()](#colletter) Convert column number to letter
  * [colLetterRange()](#colletterrange) Convert values to letters array
  * [colNumber()](#colnumber) Convert letter to number (ONE based)
  * [colNumberRange()](#colnumberrange) Convert letter range to array of numbers (ONE based)
  * [create()](#create) Create new workbook
  * [createSheet()](#createsheet)
  * [fullAddress()](#fulladdress)
  * [generateUuid()](#generateuuid) Generate UUID v4
  * [pixelsToEMU()](#pixelstoemu)
  * [rangeRelOffsets()](#rangereloffsets) Return offsets by relative address (zero based)
  * [rowIndexRange()](#rowindexrange)
  * [rowNumberRange()](#rownumberrange)
  * [setTempDir()](#settempdir) Set dir for temporary files
  * [toTimestamp()](#totimestamp) Convert value (int or string) to Excel timestamp
  * [setActiveSheet()](#setactivesheet) Set active (default) sheet by case-insensitive name
  * [addDefinedName()](#adddefinedname)
  * [addNamedRange()](#addnamedrange)
  * [addSharedString()](#addsharedstring)
  * [addStyle()](#addstyle)
  * [setDefaultFont()](#setdefaultfont) Set default font options
  * [setDefaultFontName()](#setdefaultfontname) Set default font name
  * [getDefaultFormatStyles()](#getdefaultformatstyles)
  * [setDefaultLocale()](#setdefaultlocale) Set default locale from the current environment
  * [getDefaultSheetName()](#getdefaultsheetname)
  * [getDefaultStyle()](#getdefaultstyle)
  * [setDefaultStyle()](#setdefaultstyle) Set default style
  * [getDefinedNames()](#getdefinednames)
  * [download()](#download) Download generated file to client (send to browser)
  * [getFileName()](#getfilename) Returns default filename
  * [setFileName()](#setfilename) Sets default filename for saving
  * [getHyperlinkStyle()](#gethyperlinkstyle)
  * [getImageFiles()](#getimagefiles)
  * [loadImageFile()](#loadimagefile)
  * [setLocale()](#setlocale) Set locale information
  * [makeSheet()](#makesheet)
  * [setMetaAuthor()](#setmetaauthor) Set metadata 'author'
  * [setMetaCompany()](#setmetacompany) Set metadata 'company'
  * [getMetadata()](#getmetadata) Get metadata
  * [setMetadata()](#setmetadata) Set metadata
  * [setMetaDescription()](#setmetadescription) Set metadata 'description'
  * [setMetaKeywords()](#setmetakeywords) Set metadata 'keywords'
  * [setMetaSubject()](#setmetasubject) Set metadata 'subject'
  * [setMetaTitle()](#setmetatitle) Set metadata 'title'
  * [output()](#output) Alias of download()
  * [protect()](#protect) Protect workbook
  * [removeSheet()](#removesheet) Removes sheet by index or name of sheet.
  * [isRightToLeft()](#isrighttoleft)
  * [setRightToLeft()](#setrighttoleft)
  * [save()](#save) Save generated XLSX-file
  * [getSharedStrings()](#getsharedstrings)
  * [sheet()](#sheet) Returns sheet by number or name of sheet.
  * [getSheet()](#getsheet) Alias of sheet()
  * [getSheets()](#getsheets) Returns all sheets
  * [unprotect()](#unprotect) Unprotect workbook
  * [getWriter()](#getwriter)
* [Class \avadim\FastExcelWriter\Sheet](#class-avadimfastexcelwritersheet)
  * [__construct()](#__construct) Sheet constructor
  * [setActiveCell()](#setactivecell) Set active cell
  * [addCellStyle()](#addcellstyle)
  * [addChart()](#addchart) Add chart object to the specified range of cells
  * [addDataValidation()](#adddatavalidation) Add data validation object to the specified range of cells
  * [addImage()](#addimage) Add image to the sheet
  * [addNamedRange()](#addnamedrange) Define named range
  * [addNote()](#addnote) Add note to the sheet
  * [addStyle()](#addstyle) Alias for 'addCellStyle()'
  * [allowAutoFilter()](#allowautofilter) AutoFilters should be allowed to operate when the sheet is protected
  * [allowDeleteColumns()](#allowdeletecolumns) Deleting columns should be allowed when the sheet is protected
  * [allowDeleteRows()](#allowdeleterows) Deleting rows should be allowed when the sheet is protected
  * [allowEditObjects()](#alloweditobjects) Objects are allowed to be edited when the sheet is protected
  * [allowEditScenarios()](#alloweditscenarios) Scenarios are allowed to be edited when the sheet is protected
  * [allowFormatCells()](#allowformatcells) Formatting cells should be allowed when the sheet is protected
  * [allowFormatColumns()](#allowformatcolumns) Formatting columns should be allowed when the sheet is protected
  * [allowFormatRows()](#allowformatrows) Formatting rows should be allowed when the sheet is protected
  * [allowInsertColumns()](#allowinsertcolumns) Inserting columns should be allowed when the sheet is protected
  * [allowInsertHyperlinks()](#allowinserthyperlinks) Inserting hyperlinks should be allowed when the sheet is protected
  * [allowInsertRows()](#allowinsertrows) Inserting rows should be allowed when the sheet is protected
  * [allowPivotTables()](#allowpivottables) PivotTables should be allowed to operate when the sheet is protected
  * [allowSelectCells()](#allowselectcells) Selection of any cells should be allowed when the sheet is protected
  * [allowSelectLockedCells()](#allowselectlockedcells) Selection of locked cells should be allowed when the sheet is protected
  * [allowSelectUnlockedCells()](#allowselectunlockedcells) Selection of unlocked cells should be allowed when the sheet is protected
  * [allowSort()](#allowsort) Sorting should be allowed when the sheet is protected
  * [applyBgColor()](#applybgcolor)
  * [applyBorder()](#applyborder) Sets all borders style
  * [applyBorderBottom()](#applyborderbottom)
  * [applyBorderLeft()](#applyborderleft)
  * [applyBorderRight()](#applyborderright)
  * [applyBorderTop()](#applybordertop)
  * [applyColor()](#applycolor) Alias of 'setFontColor()'
  * [applyDataValidation()](#applydatavalidation)
  * [applyFillColor()](#applyfillcolor) Alias of 'backgroundColor()'
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
  * [applyInnerBorder()](#applyinnerborder)
  * [applyNamedRange()](#applynamedrange)
  * [applyOuterBorder()](#applyouterborder)
  * [applyRowHeight()](#applyrowheight) Sets height to the current row
  * [applyRowOutlineLevel()](#applyrowoutlinelevel)
  * [applyStyle()](#applystyle)
  * [applyTextAlign()](#applytextalign)
  * [applyTextCenter()](#applytextcenter)
  * [applyTextColor()](#applytextcolor)
  * [applyTextRotation()](#applytextrotation)
  * [applyTextWrap()](#applytextwrap)
  * [applyUnlock()](#applyunlock)
  * [applyVerticalAlign()](#applyverticalalign)
  * [setAutofilter()](#setautofilter) Set auto filter
  * [beginArea()](#beginarea) Begin a new area
  * [beginOutlineLevel()](#beginoutlinelevel)
  * [setBgColor()](#setbgcolor)
  * [getBottomNodesOptions()](#getbottomnodesoptions)
  * [setBottomNodesOptions()](#setbottomnodesoptions)
  * [cell()](#cell) Select a single cell or cell range in the current row
  * [setCellStyle()](#setcellstyle) Set style for the specific cell
  * [getCharts()](#getcharts)
  * [clearAreas()](#clearareas)
  * [getColAttributes()](#getcolattributes)
  * [setColAutoWidth()](#setcolautowidth)
  * [setColDataStyle()](#setcoldatastyle) Set style of column cells (colors, formats, etc.)
  * [setColDataStyleArray()](#setcoldatastylearray) Set style of column cells (colors, formats, etc.)
  * [setColFormat()](#setcolformat) Set format of single or multiple column(s)
  * [setColFormats()](#setcolformats) Set formats of columns
  * [setColFormula()](#setcolformula) Set formula for single or multiple column(s)
  * [setColHidden()](#setcolhidden)
  * [setColMinWidth()](#setcolminwidth) Setting a minimal column's width
  * [setColMinWidths()](#setcolminwidths) Setting a multiple column's minimal width
  * [setColOptions()](#setcoloptions) Use 'setColDataStyle()' or 'setColDataStyleArray()' instead
  * [setColOutlineLevel()](#setcoloutlinelevel)
  * [setColStyle()](#setcolstyle) Set style of single or multiple column(s)
  * [setColStyleArray()](#setcolstylearray) Set style of single or multiple column(s)
  * [setColStyles()](#setcolstyles)
  * [setColVisible()](#setcolvisible) Show/hide a column
  * [setColWidth()](#setcolwidth) Set width of single or multiple column(s)
  * [setColWidthAuto()](#setcolwidthauto) Set width of single or multiple column(s)
  * [setColWidths()](#setcolwidths) Setting a multiple column's width
  * [getCurrentColId()](#getcurrentcolid)
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
  * [getDefaultStyle()](#getdefaultstyle) Returns default style
  * [setDefaultStyle()](#setdefaultstyle) Sets default style
  * [endAreas()](#endareas)
  * [endOutlineLevel()](#endoutlinelevel)
  * [setFormat()](#setformat)
  * [setFormula()](#setformula) Set a formula to the single cell or to the cell range
  * [setFreeze()](#setfreeze) Freeze rows/columns
  * [setFreezeColumns()](#setfreezecolumns) Freeze columns
  * [setFreezeRows()](#setfreezerows) Freeze rows
  * [getImages()](#getimages)
  * [getLastCell()](#getlastcell)
  * [getLastRange()](#getlastrange)
  * [makeArea()](#makearea) Make area for writing
  * [mergeCells()](#mergecells) Merge cells
  * [getMergedCells()](#getmergedcells) Returns merged cells
  * [mergeRelCells()](#mergerelcells) Merge relative cells
  * [getName()](#getname) Get sheet name
  * [isName()](#isname) Case-insensitive name checking
  * [setName()](#setname) Set sheet name
  * [getNamedRanges()](#getnamedranges) Returns named ranges with full addresses
  * [nextCell()](#nextcell)
  * [nextRow()](#nextrow) Move to the next row
  * [getNotes()](#getnotes)
  * [setOuterBorder()](#setouterborder)
  * [getOutlineLevel()](#getoutlinelevel)
  * [getPageFit()](#getpagefit)
  * [pageFitToHeight()](#pagefittoheight)
  * [getPageFitToHeight()](#getpagefittoheight)
  * [pageFitToWidth()](#pagefittowidth)
  * [getPageFitToWidth()](#getpagefittowidth)
  * [pageLandscape()](#pagelandscape) Set page orientation as Landscape
  * [pageMarginBottom()](#pagemarginbottom) Bottom Page Margin in mm|cm|in
  * [pageMarginFooter()](#pagemarginfooter) Footer Page Margin in mm|cm|in
  * [pageMarginHeader()](#pagemarginheader) Header Page Margin in mm|cm|in
  * [pageMarginLeft()](#pagemarginleft) Left Page Margin in mm|cm|in
  * [pageMarginRight()](#pagemarginright) Right page margin in mm|cm|in
  * [pageMargins()](#pagemargins) Page margins for a sheet or a custom sheet view in mm|cm|in
  * [getPageMargins()](#getpagemargins)
  * [setPageMargins()](#setpagemargins)
  * [pageMarginTop()](#pagemargintop) Top Page Margin in mm|cm|in
  * [setPageOptions()](#setpageoptions)
  * [getPageOrientation()](#getpageorientation)
  * [pageOrientationLandscape()](#pageorientationlandscape) Set page orientation as Landscape, alias of pageLandscape()
  * [pageOrientationPortrait()](#pageorientationportrait) Set page orientation as Portrait, alias of pagePortrait()
  * [pagePaperHeight()](#pagepaperheight) Height of custom paper as a number followed by a unit identifier mm|cm|in (ex: 297mm, 11in)
  * [pagePaperSize()](#pagepapersize) Set Paper size (when paperHeight and paperWidth are specified, paperSize should be ignored)
  * [pagePaperWidth()](#pagepaperwidth) Width of custom paper as a number followed by a unit identifier mm|cm|in (ex: 21cm, 8.5in)
  * [pagePortrait()](#pageportrait) Set page orientation as Portrait
  * [getPageSetup()](#getpagesetup)
  * [setPageSetup()](#setpagesetup)
  * [setPrintArea()](#setprintarea)
  * [setPrintGridlines()](#setprintgridlines)
  * [setPrintLeftColumns()](#setprintleftcolumns)
  * [setPrintTitles()](#setprinttitles)
  * [setPrintTopRows()](#setprinttoprows)
  * [protect()](#protect) Protect sheet
  * [isRightToLeft()](#isrighttoleft)
  * [setRowDataStyle()](#setrowdatastyle) Style are applied only to non-empty cells in a row (or row range)
  * [setRowDataStyleArray()](#setrowdatastylearray) Styles are applied only to non-empty cells in a rows
  * [setRowHeight()](#setrowheight) Height of a specific row
  * [setRowHeights()](#setrowheights) Multiple rows height
  * [setRowHidden()](#setrowhidden) Hide a specific row
  * [setRowOutlineLevel()](#setrowoutlinelevel)
  * [setRowStyle()](#setrowstyle) Style are applied to the entire sheet row (even if it is empty)
  * [setRowStyleArray()](#setrowstylearray) Styles are applied to the entire sheet row (even if it is empty)
  * [setRowVisible()](#setrowvisible) Hide/show a specific row
  * [skipRow()](#skiprow) Skip rows
  * [setStyle()](#setstyle) Alias for 'setCellStyle()'
  * [setTabColor()](#settabcolor) Set color for the sheet tab
  * [setTopLeftCell()](#settopleftcell) Set top left cell for writing
  * [unprotect()](#unprotect) Unprotect sheet
  * [setValue()](#setvalue) Set a value to the single cell or to the cell range
  * [withLastCell()](#withlastcell) Select last written cell for applying
  * [withLastRow()](#withlastrow) Select last written row for applying
  * [withRange()](#withrange) Select custom range for applying
  * [writeAreas()](#writeareas)
  * [writeArray()](#writearray) Write values from two-dimensional array
  * [writeArrayTo()](#writearrayto) Write 2d array form the specified cell
  * [writeCell()](#writecell) Write value to the current cell and move pointer to the next cell in the row
  * [writeHeader()](#writeheader)
  * [writeRow()](#writerow) Write values to the current row
  * [writeTo()](#writeto) Write value to the specified cell and move pointer to the next cell in the row

---

## Class \avadim\FastExcelWriter\Excel

### __construct()

---

```php
public __construct(?array $options = [])
```
_Excel constructor_

#### Parameters

`array|null $options` -- Optional parameters: ['temp_dir' => ..., 'temp_prefix' => ..., 'auto_convert_number' => ..., 'shared_string' => ...]

---

### cellAddress()

---

```php
public static cellAddress(int $rowNumber, int $colNumber, ?bool $absolute, ?bool $absoluteRow): string
```
_Create cell address by row and col numbers_

#### Parameters

`int $rowNumber` -- ONE based

`int $colNumber` -- ONE based

`bool|null $absolute`

`bool|null $absoluteRow`

---

### colIndex()

---

```php
public static colIndex(string $colLetter): int
```
_Convert letter to index (ZERO based)_

#### Parameters

`string $colLetter`

---

### colIndexRange()

---

```php
public static colIndexRange($colLetter): array
```
_Convert letter range to array of numbers (ZERO based)_

#### Parameters

`string|int|array $colLetter` -- e.g.: 'B', 2, 'C:F', ['A', 'B', 'C']

---

### colKeysToIndexes()

---

```php
public static colKeysToIndexes(array $data): array
```


#### Parameters

`array $data`

---

### colKeysToLetters()

---

```php
public static colKeysToLetters(array $data): array
```


#### Parameters

`array $data`

---

### colKeysToNumbers()

---

```php
public static colKeysToNumbers(array $data): array
```


#### Parameters

`array $data`

---

### colLetter()

---

```php
public static colLetter(int $colNumber): string
```
_Convert column number to letter_

#### Parameters

`int $colNumber` -- ONE based

---

### colLetterRange()

---

```php
public static colLetterRange($colKeys, ?int $baseNum): array
```
_Convert values to letters array_

#### Parameters

`array|string $colKeys`

`int|null $baseNum` -- 0 or 1

---

### colNumber()

---

```php
public static colNumber(string $colLetter): int
```
_Convert letter to number (ONE based)_

#### Parameters

`string $colLetter`

---

### colNumberRange()

---

```php
public static colNumberRange($colLetter): array
```
_Convert letter range to array of numbers (ONE based)_

#### Parameters

`string|int|array $colLetter` -- e.g.: 'B', 2, 'C:F', ['A', 'B', 'C']

---

### create()

---

```php
public static create($sheets, ?array $options = []): Excel
```
_Create new workbook_

#### Parameters

`array|string|null $sheets` -- Name of sheet or array of names

`array|null $options` -- Options

---

### createSheet()

---

```php
public static createSheet(string $sheetName): Sheet
```


#### Parameters

`string $sheetName`

---

### fullAddress()

---

```php
public static fullAddress(string $sheetName, string $address, ?bool $force): string
```


#### Parameters

`string $sheetName`

`string $address`

`bool|null $force`

---

### generateUuid()

---

```php
public static generateUuid(): string
```
_Generate UUID v4_

#### Parameters

_None_

---

### pixelsToEMU()

---

```php
public static pixelsToEMU($pixels): float|int
```


#### Parameters

`$pixels`

---

### rangeRelOffsets()

---

```php
public static rangeRelOffsets(string $relAddress): array
```
_Return offsets by relative address (zero based)_

#### Parameters

`string $relAddress`

---

### rowIndexRange()

---

```php
public static rowIndexRange($rowRange): array
```


#### Parameters

`int|string|array $rowRange`

---

### rowNumberRange()

---

```php
public static rowNumberRange($rowRange): array
```


#### Parameters

`int|string|array $rowRange`

---

### setTempDir()

---

```php
public static setTempDir($tempDir)
```
_Set dir for temporary files_

#### Parameters

`$tempDir`

---

### toTimestamp()

---

```php
public static toTimestamp($value): float|bool
```
_Convert value (int or string) to Excel timestamp_

#### Parameters

`int|string $value`

---

### setActiveSheet()

---

```php
public setActiveSheet(string $name): Excel
```
_Set active (default) sheet by case-insensitive name_

#### Parameters

`string $name`

---

### addDefinedName()

---

```php
public addDefinedName(string $name, string $range, ?array $attributes = []): Excel
```


#### Parameters

`string $name`

`string $range`

`array|null $attributes`

---

### addNamedRange()

---

```php
public addNamedRange(string $range, string $name): Excel
```


#### Parameters

`string $range`

`string $name`

---

### addSharedString()

---

```php
public addSharedString(string $string, ?bool $richText): int
```


#### Parameters

`string $string`

`bool|null $richText`

---

### addStyle()

---

```php
public addStyle($cellStyle, &$resultStyle): int
```


#### Parameters

`$cellStyle`

`$resultStyle`

---

### setDefaultFont()

---

```php
public setDefaultFont(array $fontOptions): Excel
```
_Set default font options_

#### Parameters

`array $fontOptions`

---

### setDefaultFontName()

---

```php
public setDefaultFontName(string $fontName): Excel
```
_Set default font name_

#### Parameters

`string $fontName`

---

### getDefaultFormatStyles()

---

```php
public getDefaultFormatStyles(): array
```


#### Parameters

_None_

---

### setDefaultLocale()

---

```php
public setDefaultLocale()
```
_Set default locale from the current environment_

#### Parameters

_None_

---

### getDefaultSheetName()

---

```php
public getDefaultSheetName(): string
```


#### Parameters

_None_

---

### getDefaultStyle()

---

```php
public getDefaultStyle(): array
```


#### Parameters

_None_

---

### setDefaultStyle()

---

```php
public setDefaultStyle(array $style): Excel
```
_Set default style_

#### Parameters

`array $style`

---

### getDefinedNames()

---

```php
public getDefinedNames(): array
```


#### Parameters

_None_

---

### download()

---

```php
public download(?string $name)
```
_Download generated file to client (send to browser)_

#### Parameters

`string|null $name`

---

### getFileName()

---

```php
public getFileName(): string
```
_Returns default filename_

#### Parameters

_None_

---

### setFileName()

---

```php
public setFileName(string $fileName): Excel
```
_Sets default filename for saving_

#### Parameters

`string $fileName`

---

### getHyperlinkStyle()

---

```php
public getHyperlinkStyle(): array
```


#### Parameters

_None_

---

### getImageFiles()

---

```php
public getImageFiles(): array
```


#### Parameters

_None_

---

### loadImageFile()

---

```php
public loadImageFile(string $imageFile): ?array
```


#### Parameters

`string $imageFile`

---

### setLocale()

---

```php
public setLocale(string $locale, ?string $dir): Excel
```
_Set locale information_

#### Parameters

`string $locale`

`string|null $dir`

---

### makeSheet()

---

```php
public makeSheet(?string $sheetName): Sheet
```


#### Parameters

`string|null $sheetName`

---

### setMetaAuthor()

---

```php
public setMetaAuthor(?string $author): Excel
```
_Set metadata 'author'_

#### Parameters

`string|null $author`

---

### setMetaCompany()

---

```php
public setMetaCompany(?string $company): Excel
```
_Set metadata 'company'_

#### Parameters

`string|null $company`

---

### getMetadata()

---

```php
public getMetadata($key): mixed
```
_Get metadata_

#### Parameters

`null $key`

---

### setMetadata()

---

```php
public setMetadata($key, $value): Excel
```
_Set metadata_

#### Parameters

`$key`

`$value`

---

### setMetaDescription()

---

```php
public setMetaDescription(?string $description): Excel
```
_Set metadata 'description'_

#### Parameters

`string|null $description`

---

### setMetaKeywords()

---

```php
public setMetaKeywords($keywords): Excel
```
_Set metadata 'keywords'_

#### Parameters

`mixed $keywords`

---

### setMetaSubject()

---

```php
public setMetaSubject(?string $subject): Excel
```
_Set metadata 'subject'_

#### Parameters

`string|null $subject`

---

### setMetaTitle()

---

```php
public setMetaTitle(?string $title): Excel
```
_Set metadata 'title'_

#### Parameters

`string|null $title`

---

### output()

---

```php
public output(?string $name): void
```
_Alias of download()_

#### Parameters

`string|null $name`

---

### protect()

---

```php
public protect(?string $password): Excel
```
_Protect workbook_

#### Parameters

`string|null $password`

---

### removeSheet()

---

```php
public removeSheet($index): Excel
```
_Removes sheet by index or name of sheet.Removes the first sheet of index omitted_

#### Parameters

`int|string|null $index`

---

### isRightToLeft()

---

```php
public isRightToLeft(): bool
```


#### Parameters

_None_

---

### setRightToLeft()

---

```php
public setRightToLeft(bool $isRightToLeft)
```


#### Parameters

`bool $isRightToLeft`

---

### save()

---

```php
public save(?string $fileName, ?bool $overWrite): bool
```
_Save generated XLSX-file_

#### Parameters

`string|null $fileName`

`bool|null $overWrite`

---

### getSharedStrings()

---

```php
public getSharedStrings(): array
```


#### Parameters

_None_

---

### sheet()

---

```php
public sheet($index): ?avadim\FastExcelWriter\Sheet
```
_Returns sheet by number or name of sheet.Return the first sheet if number or name omitted_

#### Parameters

`int|string|null $index` -- - number or name of sheet

---

### getSheet()

---

```php
public getSheet($index): ?avadim\FastExcelWriter\Sheet
```
_Alias of sheet()_

#### Parameters

`int|string|null $index` -- - number or name of sheet

---

### getSheets()

---

```php
public getSheets(): array
```
_Returns all sheets_

#### Parameters

_None_

---

### unprotect()

---

```php
public unprotect(): Excel
```
_Unprotect workbook_

#### Parameters

_None_

---

### getWriter()

---

```php
public getWriter(): Writer
```


#### Parameters

_None_

---

## Class \avadim\FastExcelWriter\Sheet

### __construct()

---

```php
public __construct(string $sheetName)
```
_Sheet constructor_

#### Parameters

`string $sheetName`

---

### setActiveCell()

---

```php
public setActiveCell($cellAddress): Sheet
```
_Set active cell_

#### Parameters

`$cellAddress`

---

### addCellStyle()

---

```php
public addCellStyle(string $cellAddr, array $style): Sheet
```


#### Parameters

`string $cellAddr`

`array $style`

---

### addChart()

---

```php
public addChart(string $range, avadim\FastExcelWriter\Charts\Chart $chart): Sheet
```
_Add chart object to the specified range of cells_

#### Parameters

`string $range` -- Set the position where the chart should appear in the worksheet

`Chart $chart` -- Chart object

---

### addDataValidation()

---

```php
public addDataValidation(string $range, avadim\FastExcelWriter\DataValidation\DataValidation $validation): Sheet
```
_Add data validation object to the specified range of cells_

#### Parameters

`string $range`

`DataValidation $validation`

---

### addImage()

---

```php
public addImage(string $cell, string $imageFile, ?array $imageStyle = []): Sheet
```
_Add image to the sheet_

#### Parameters

`string $cell`

`string $imageFile`

`array|null $imageStyle`

---

### addNamedRange()

---

```php
public addNamedRange(string $range, string $name): Sheet
```
_Define named range_

#### Parameters

`string $range`

`string $name`

---

### addNote()

---

```php
public addNote($cell, $comment, array $noteStyle = []): Sheet
```
_Add note to the sheet_

#### Parameters

`string|mixed $cell`

`string|array|null $comment`

`array $noteStyle`

---

### addStyle()

---

```php
public addStyle(string $cellAddr, array $style): Sheet
```
_Alias for 'addCellStyle()'_

#### Parameters

`string $cellAddr`

`array $style`

---

### allowAutoFilter()

---

```php
public allowAutoFilter(?bool $allow): Sheet
```
_AutoFilters should be allowed to operate when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowDeleteColumns()

---

```php
public allowDeleteColumns(?bool $allow): Sheet
```
_Deleting columns should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowDeleteRows()

---

```php
public allowDeleteRows(?bool $allow): Sheet
```
_Deleting rows should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowEditObjects()

---

```php
public allowEditObjects(?bool $allow): Sheet
```
_Objects are allowed to be edited when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowEditScenarios()

---

```php
public allowEditScenarios(?bool $allow): Sheet
```
_Scenarios are allowed to be edited when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowFormatCells()

---

```php
public allowFormatCells(?bool $allow): Sheet
```
_Formatting cells should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowFormatColumns()

---

```php
public allowFormatColumns(?bool $allow): Sheet
```
_Formatting columns should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowFormatRows()

---

```php
public allowFormatRows(?bool $allow): Sheet
```
_Formatting rows should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowInsertColumns()

---

```php
public allowInsertColumns(?bool $allow): Sheet
```
_Inserting columns should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowInsertHyperlinks()

---

```php
public allowInsertHyperlinks(?bool $allow): Sheet
```
_Inserting hyperlinks should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowInsertRows()

---

```php
public allowInsertRows(?bool $allow): Sheet
```
_Inserting rows should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowPivotTables()

---

```php
public allowPivotTables(?bool $allow): Sheet
```
_PivotTables should be allowed to operate when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowSelectCells()

---

```php
public allowSelectCells(?bool $allow): Sheet
```
_Selection of any cells should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowSelectLockedCells()

---

```php
public allowSelectLockedCells(?bool $allow): Sheet
```
_Selection of locked cells should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowSelectUnlockedCells()

---

```php
public allowSelectUnlockedCells(?bool $allow): Sheet
```
_Selection of unlocked cells should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### allowSort()

---

```php
public allowSort(?bool $allow): Sheet
```
_Sorting should be allowed when the sheet is protected_

#### Parameters

`bool|null $allow`

---

### applyBgColor()

---

```php
public applyBgColor(string $color): Sheet
```


#### Parameters

`string $color`

---

### applyBorder()

---

```php
public applyBorder(string $style, ?string $color): Sheet
```
_Sets all borders style_

#### Parameters

`string $style`

`string|null $color`

---

### applyBorderBottom()

---

```php
public applyBorderBottom(string $style, ?string $color): Sheet
```


#### Parameters

`string $style`

`string|null $color`

---

### applyBorderLeft()

---

```php
public applyBorderLeft(string $style, ?string $color): Sheet
```


#### Parameters

`string $style`

`string|null $color`

---

### applyBorderRight()

---

```php
public applyBorderRight(string $style, ?string $color): Sheet
```


#### Parameters

`string $style`

`string|null $color`

---

### applyBorderTop()

---

```php
public applyBorderTop(string $style, ?string $color): Sheet
```


#### Parameters

`string $style`

`string|null $color`

---

### applyColor()

---

```php
public applyColor(string $color): Sheet
```
_Alias of 'setFontColor()'_

#### Parameters

`string $color`

---

### applyDataValidation()

---

```php
public applyDataValidation(avadim\FastExcelWriter\DataValidation\DataValidation $validation): Sheet
```


#### Parameters

`DataValidation $validation`

---

### applyFillColor()

---

```php
public applyFillColor(string $color): Sheet
```
_Alias of 'backgroundColor()'_

#### Parameters

`string $color`

---

### applyFont()

---

```php
public applyFont(string $fontName, ?int $fontSize, ?string $fontStyle, ?string $fontColor): Sheet
```


#### Parameters

`string $fontName`

`int|null $fontSize`

`string|null $fontStyle`

`string|null $fontColor`

---

### applyFontColor()

---

```php
public applyFontColor(string $fontColor): Sheet
```


#### Parameters

`string $fontColor`

---

### applyFontName()

---

```php
public applyFontName(string $fontName): Sheet
```


#### Parameters

`string $fontName`

---

### applyFontSize()

---

```php
public applyFontSize(float $fontSize): Sheet
```


#### Parameters

`float $fontSize`

---

### applyFontStyle()

---

```php
public applyFontStyle(string $fontStyle): Sheet
```


#### Parameters

`string $fontStyle`

---

### applyFontStyleBold()

---

```php
public applyFontStyleBold(): Sheet
```


#### Parameters

_None_

---

### applyFontStyleItalic()

---

```php
public applyFontStyleItalic(): Sheet
```


#### Parameters

_None_

---

### applyFontStyleStrikethrough()

---

```php
public applyFontStyleStrikethrough(): Sheet
```


#### Parameters

_None_

---

### applyFontStyleUnderline()

---

```php
public applyFontStyleUnderline(?bool $double): Sheet
```


#### Parameters

`bool|null $double`

---

### applyFormat()

---

```php
public applyFormat($format): Sheet
```


#### Parameters

`string|array $format`

---

### applyHide()

---

```php
public applyHide(?bool $hide): Sheet
```


#### Parameters

`$hide`

---

### applyInnerBorder()

---

```php
public applyInnerBorder(string $style, ?string $color): Sheet
```


#### Parameters

`string $style`

`string|null $color`

---

### applyNamedRange()

---

```php
public applyNamedRange(string $name): Sheet
```


#### Parameters

`string $name`

---

### applyOuterBorder()

---

```php
public applyOuterBorder(string $style, ?string $color): Sheet
```


#### Parameters

`string $style`

`string|null $color`

---

### applyRowHeight()

---

```php
public applyRowHeight(float $height): Sheet
```
_Sets height to the current row_

#### Parameters

`float $height`

---

### applyRowOutlineLevel()

---

```php
public applyRowOutlineLevel(int $outlineLevel): Sheet
```


#### Parameters

`int $outlineLevel`

---

### applyStyle()

---

```php
public applyStyle(array $style): Sheet
```


#### Parameters

`array $style`

---

### applyTextAlign()

---

```php
public applyTextAlign(string $textAlign, ?string $verticalAlign): Sheet
```


#### Parameters

`string $textAlign`

`string|null $verticalAlign`

---

### applyTextCenter()

---

```php
public applyTextCenter(): Sheet
```


#### Parameters

_None_

---

### applyTextColor()

---

```php
public applyTextColor(string $color): Sheet
```


#### Parameters

`string $color`

---

### applyTextRotation()

---

```php
public applyTextRotation(int $degrees): Sheet
```


#### Parameters

`int $degrees`

---

### applyTextWrap()

---

```php
public applyTextWrap(?bool $textWrap): Sheet
```


#### Parameters

`bool|null $textWrap`

---

### applyUnlock()

---

```php
public applyUnlock(?bool $unlock): Sheet
```


#### Parameters

`bool $unlock`

---

### applyVerticalAlign()

---

```php
public applyVerticalAlign(string $verticalAlign): Sheet
```


#### Parameters

`string $verticalAlign`

---

### setAutofilter()

---

```php
public setAutofilter(?int $row, ?int $col): Sheet
```
_Set auto filter_

#### Parameters

`int|null $row`

`int|null $col`

---

### beginArea()

---

```php
public beginArea(?string $cellAddress): Area
```
_Begin a new area_

#### Parameters

`string|null $cellAddress` -- Upper left cell of area

---

### beginOutlineLevel()

---

```php
public beginOutlineLevel(?bool $collapsed): Sheet
```


#### Parameters

`$collapsed`

---

### setBgColor()

---

```php
public setBgColor(string $cellAddr, string $color): Sheet
```


#### Parameters

`string $cellAddr`

`string $color`

---

### getBottomNodesOptions()

---

```php
public getBottomNodesOptions(): array
```


#### Parameters

_None_

---

### setBottomNodesOptions()

---

```php
public setBottomNodesOptions(string $node, array $options): Sheet
```


#### Parameters

`string $node`

`array $options`

---

### cell()

---

```php
public cell($cellAddress): Sheet
```
_Select a single cell or cell range in the current row_

_$cellAddress formats:'B5''B5:C7'['col' => 2, 'row' => 5][2, 5]_

#### Parameters

`string|array $cellAddress`

---

### setCellStyle()

---

```php
public setCellStyle(string $cellAddress, $style, ?bool $mergeStyles): Sheet
```
_Set style for the specific cell_

#### Parameters

`string $cellAddress` -- Cell address

`mixed $style` -- Style array or object

`bool|null $mergeStyles` -- True - merge style with previous style for this cell (if exists)

---

### getCharts()

---

```php
public getCharts(): array
```


#### Parameters

_None_

---

### clearAreas()

---

```php
public clearAreas(): Sheet
```


#### Parameters

_None_

---

### getColAttributes()

---

```php
public getColAttributes(): array
```


#### Parameters

_None_

---

### setColAutoWidth()

---

```php
public setColAutoWidth($col): Sheet
```


#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

---

### setColDataStyle()

---

```php
public setColDataStyle($colRange, array $colStyle): Sheet
```
_Set style of column cells (colors, formats, etc.)_

_Styles are applied only to non-empty cells in a column and only take effect starting with the current row_

#### Parameters

`int|string|array $colRange`

`array $colStyle`

---

### setColDataStyleArray()

---

```php
public setColDataStyleArray(array $colStyles): Sheet
```
_Set style of column cells (colors, formats, etc.)_

_Styles are applied only to non-empty cells in a column and only take effect starting with the current row_

#### Parameters

`array $colStyles`

---

### setColFormat()

---

```php
public setColFormat($col, $format): Sheet
```
_Set format of single or multiple column(s)_

#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

`mixed $format`

---

### setColFormats()

---

```php
public setColFormats(array $formats): Sheet
```
_Set formats of columns_

#### Parameters

`array $formats`

---

### setColFormula()

---

```php
public setColFormula($col, string $formula): Sheet
```
_Set formula for single or multiple column(s)_

#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

`string $formula`

---

### setColHidden()

---

```php
public setColHidden($col): Sheet
```


#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

---

### setColMinWidth()

---

```php
public setColMinWidth($col, $width): Sheet
```
_Setting a minimal column's width_

#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

`int|float|string $width`

---

### setColMinWidths()

---

```php
public setColMinWidths(array $widths): Sheet
```
_Setting a multiple column's minimal width_

#### Parameters

`array $widths`

---

### setColOptions()

---

```php
public setColOptions($arg1, ?array $arg2): Sheet
```
_Use 'setColDataStyle()' or 'setColDataStyleArray()' instead_

#### Parameters

`$arg1`

`$arg2`

---

### setColOutlineLevel()

---

```php
public setColOutlineLevel($col, int $outlineLevel): Sheet
```


#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

`int $outlineLevel`

---

### setColStyle()

---

```php
public setColStyle($colRange, $style): Sheet
```
_Set style of single or multiple column(s)_

_Styles are applied to the entire sheet column(s) (even if it is empty)_

#### Parameters

`int|string|array $colRange` -- Column number or column letter (or array of these)

`mixed $style`

---

### setColStyleArray()

---

```php
public setColStyleArray(array $colStyles): Sheet
```
_Set style of single or multiple column(s)_

_Styles are applied to the entire sheet column(s) (even if it is empty)_

#### Parameters

`array $colStyles`

---

### setColStyles()

---

```php
public setColStyles($arg1, ?array $arg2): Sheet
```


#### Parameters

`$arg1`

`array|null $arg2`

---

### setColVisible()

---

```php
public setColVisible($col, bool $val): Sheet
```
_Show/hide a column_

#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

`bool $val`

---

### setColWidth()

---

```php
public setColWidth($col, $width, ?bool $min): Sheet
```
_Set width of single or multiple column(s)_

#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

`int|float|string $width`

`bool|null $min`

---

### setColWidthAuto()

---

```php
public setColWidthAuto($col): Sheet
```
_Set width of single or multiple column(s)_

#### Parameters

`int|string|array $col` -- Column number or column letter (or array of these)

---

### setColWidths()

---

```php
public setColWidths(array $widths, ?bool $min): Sheet
```
_Setting a multiple column's width_

#### Parameters

`array $widths`

`bool|null $min`

---

### getCurrentColId()

---

```php
public getCurrentColId(): int
```


#### Parameters

_None_

---

### getCurrentRowId()

---

```php
public getCurrentRowId(): int
```


#### Parameters

_None_

---

### getDataValidations()

---

```php
public getDataValidations(): array
```


#### Parameters

_None_

---

### setDefaultFont()

---

```php
public setDefaultFont($font): Sheet
```


#### Parameters

`string|array $font`

---

### setDefaultFontColor()

---

```php
public setDefaultFontColor(string $fontColor): Sheet
```


#### Parameters

`string $fontColor`

---

### setDefaultFontName()

---

```php
public setDefaultFontName(string $fontName): Sheet
```


#### Parameters

`string $fontName`

---

### setDefaultFontSize()

---

```php
public setDefaultFontSize(int $fontSize): Sheet
```


#### Parameters

`int $fontSize`

---

### setDefaultFontStyle()

---

```php
public setDefaultFontStyle(string $fontStyle): Sheet
```


#### Parameters

`string $fontStyle`

---

### setDefaultFontStyleBold()

---

```php
public setDefaultFontStyleBold(): Sheet
```


#### Parameters

_None_

---

### setDefaultFontStyleItalic()

---

```php
public setDefaultFontStyleItalic(): Sheet
```


#### Parameters

_None_

---

### setDefaultFontStyleStrikethrough()

---

```php
public setDefaultFontStyleStrikethrough(): Sheet
```


#### Parameters

_None_

---

### setDefaultFontStyleUnderline()

---

```php
public setDefaultFontStyleUnderline(?bool $double): Sheet
```


#### Parameters

`bool|null $double`

---

### getDefaultStyle()

---

```php
public getDefaultStyle(): array
```
_Returns default style_

#### Parameters

_None_

---

### setDefaultStyle()

---

```php
public setDefaultStyle(array $style): Sheet
```
_Sets default style_

#### Parameters

`array $style`

---

### endAreas()

---

```php
public endAreas(): Sheet
```


#### Parameters

_None_

---

### endOutlineLevel()

---

```php
public endOutlineLevel(): Sheet
```


#### Parameters

_None_

---

### setFormat()

---

```php
public setFormat(string $cellAddr, string $format): Sheet
```


#### Parameters

`string $cellAddr`

`string $format`

---

### setFormula()

---

```php
public setFormula($cellAddress, $value, ?array $styles): Sheet
```
_Set a formula to the single cell or to the cell range_

_$cellAddress formats:'B5''B5:C7'['col' => 2, 'row' => 5][2, 5]_

#### Parameters

`string|array $cellAddress`

`mixed $value`

`array|null $styles`

---

### setFreeze()

---

```php
public setFreeze($freezeRows, $freezeColumns): Sheet
```
_Freeze rows/columns_

#### Parameters

`mixed $freezeRows`

`mixed $freezeColumns`

---

### setFreezeColumns()

---

```php
public setFreezeColumns(int $freezeColumns): Sheet
```
_Freeze columns_

#### Parameters

`int $freezeColumns` -- Number columns to freeze

---

### setFreezeRows()

---

```php
public setFreezeRows(int $freezeRows): Sheet
```
_Freeze rows_

#### Parameters

`int $freezeRows` -- Number rows to freeze

---

### getImages()

---

```php
public getImages(): array
```


#### Parameters

_None_

---

### getLastCell()

---

```php
public getLastCell(?bool $absolute): string
```


#### Parameters

`bool|null $absolute`

---

### getLastRange()

---

```php
public getLastRange(?bool $absolute): string
```


#### Parameters

`bool|null $absolute`

---

### makeArea()

---

```php
public makeArea(string $range): Area
```
_Make area for writing_

#### Parameters

`string $range` -- A1:Z9 or R1C1:R9C28

---

### mergeCells()

---

```php
public mergeCells($rangeSet, ?int $actionMode): Sheet
```
_Merge cells_

#### Parameters

`array|string|int $rangeSet`

`int|null $actionMode` -- Action in case of intersection 0 - exception 1 - replace 2 - keep -1 - skip intersection check

---

### getMergedCells()

---

```php
public getMergedCells(): array
```
_Returns merged cells_

#### Parameters

_None_

---

### mergeRelCells()

---

```php
public mergeRelCells($rangeSet): Sheet
```
_Merge relative cells_

#### Parameters

`array|string|int $rangeSet`

---

### getName()

---

```php
public getName(): string
```
_Get sheet name_

#### Parameters

_None_

---

### isName()

---

```php
public isName(string $name): bool
```
_Case-insensitive name checking_

#### Parameters

`string $name`

---

### setName()

---

```php
public setName(string $sheetName): Sheet
```
_Set sheet name_

#### Parameters

`string $sheetName`

---

### getNamedRanges()

---

```php
public getNamedRanges(): array
```
_Returns named ranges with full addresses_

#### Parameters

_None_

---

### nextCell()

---

```php
public nextCell(): Sheet
```


#### Parameters

_None_

---

### nextRow()

---

```php
public nextRow(?array $style = []): Sheet
```
_Move to the next row_

#### Parameters

`array|null $style`

---

### getNotes()

---

```php
public getNotes(): array
```


#### Parameters

_None_

---

### setOuterBorder()

---

```php
public setOuterBorder(string $range, $style): Sheet
```


#### Parameters

`string $range`

`string|array $style`

---

### getOutlineLevel()

---

```php
public getOutlineLevel(): int
```


#### Parameters

_None_

---

### getPageFit()

---

```php
public getPageFit(): bool
```


#### Parameters

_None_

---

### pageFitToHeight()

---

```php
public pageFitToHeight($numPage): Sheet
```


#### Parameters

`int|string|null $numPage`

---

### getPageFitToHeight()

---

```php
public getPageFitToHeight(): int
```


#### Parameters

_None_

---

### pageFitToWidth()

---

```php
public pageFitToWidth($numPage): Sheet
```


#### Parameters

`int|string|null $numPage`

---

### getPageFitToWidth()

---

```php
public getPageFitToWidth(): int
```


#### Parameters

_None_

---

### pageLandscape()

---

```php
public pageLandscape(): Sheet
```
_Set page orientation as Landscape_

#### Parameters

_None_

---

### pageMarginBottom()

---

```php
public pageMarginBottom($value): Sheet
```
_Bottom Page Margin in mm|cm|in_

#### Parameters

`string|float $value`

---

### pageMarginFooter()

---

```php
public pageMarginFooter($value): Sheet
```
_Footer Page Margin in mm|cm|in_

#### Parameters

`string|float $value`

---

### pageMarginHeader()

---

```php
public pageMarginHeader($value): Sheet
```
_Header Page Margin in mm|cm|in_

#### Parameters

`string|float $value`

---

### pageMarginLeft()

---

```php
public pageMarginLeft($value): Sheet
```
_Left Page Margin in mm|cm|in_

#### Parameters

`string|float $value`

---

### pageMarginRight()

---

```php
public pageMarginRight($value): Sheet
```
_Right page margin in mm|cm|in_

#### Parameters

`string|float $value`

---

### pageMargins()

---

```php
public pageMargins(array $margins): Sheet
```
_Page margins for a sheet or a custom sheet view in mm|cm|in_

#### Parameters

`array $margins`

---

### getPageMargins()

---

```php
public getPageMargins(): array
```


#### Parameters

_None_

---

### setPageMargins()

---

```php
public setPageMargins(array $margins): Sheet
```


#### Parameters

`array $margins`

---

### pageMarginTop()

---

```php
public pageMarginTop($value): Sheet
```
_Top Page Margin in mm|cm|in_

#### Parameters

`string|float $value`

---

### setPageOptions()

---

```php
public setPageOptions(string $option, $value): Sheet
```


#### Parameters

`string $option`

`mixed $value`

---

### getPageOrientation()

---

```php
public getPageOrientation(): string
```


#### Parameters

_None_

---

### pageOrientationLandscape()

---

```php
public pageOrientationLandscape(): Sheet
```
_Set page orientation as Landscape, alias of pageLandscape()_

#### Parameters

_None_

---

### pageOrientationPortrait()

---

```php
public pageOrientationPortrait(): Sheet
```
_Set page orientation as Portrait, alias of pagePortrait()_

#### Parameters

_None_

---

### pagePaperHeight()

---

```php
public pagePaperHeight($paperHeight): Sheet
```
_Height of custom paper as a number followed by a unit identifier mm|cm|in (ex: 297mm, 11in)_

#### Parameters

`string|float|int $paperHeight`

---

### pagePaperSize()

---

```php
public pagePaperSize(int $paperSize): Sheet
```
_Set Paper size (when paperHeight and paperWidth are specified, paperSize should be ignored)_

#### Parameters

`int $paperSize`

---

### pagePaperWidth()

---

```php
public pagePaperWidth($paperWidth): Sheet
```
_Width of custom paper as a number followed by a unit identifier mm|cm|in (ex: 21cm, 8.5in)_

#### Parameters

`string|float|int $paperWidth`

---

### pagePortrait()

---

```php
public pagePortrait(): Sheet
```
_Set page orientation as Portrait_

#### Parameters

_None_

---

### getPageSetup()

---

```php
public getPageSetup(): array
```


#### Parameters

_None_

---

### setPageSetup()

---

```php
public setPageSetup(array $options): Sheet
```


#### Parameters

`array $options`

---

### setPrintArea()

---

```php
public setPrintArea(string $range): Sheet
```


#### Parameters

`string $range`

---

### setPrintGridlines()

---

```php
public setPrintGridlines(bool $bool): Sheet
```


#### Parameters

`bool $bool`

---

### setPrintLeftColumns()

---

```php
public setPrintLeftColumns(string $range): Sheet
```


#### Parameters

`string $range`

---

### setPrintTitles()

---

```php
public setPrintTitles(?string $rowsAtTop, ?string $colsAtLeft): Sheet
```


#### Parameters

`string|null $rowsAtTop`

`string|null $colsAtLeft`

---

### setPrintTopRows()

---

```php
public setPrintTopRows(string $range): Sheet
```


#### Parameters

`string $range`

---

### protect()

---

```php
public protect(?string $password): Sheet
```
_Protect sheet_

#### Parameters

`string|null $password`

---

### isRightToLeft()

---

```php
public isRightToLeft(): bool
```


#### Parameters

_None_

---

### setRowDataStyle()

---

```php
public setRowDataStyle($rowRange, array $style): Sheet
```
_Style are applied only to non-empty cells in a row (or row range)_

#### Parameters

`int|string|array $rowRange`

`array $style`

---

### setRowDataStyleArray()

---

```php
public setRowDataStyleArray(array $rowStyles): Sheet
```
_Styles are applied only to non-empty cells in a rows_

#### Parameters

`array $rowStyles`

---

### setRowHeight()

---

```php
public setRowHeight($rowNum, $height): Sheet
```
_Height of a specific row_

#### Parameters

`$rowNum`

`$height`

---

### setRowHeights()

---

```php
public setRowHeights(array $heights): Sheet
```
_Multiple rows height_

#### Parameters

`array $heights`

---

### setRowHidden()

---

```php
public setRowHidden($rowNum): Sheet
```
_Hide a specific row_

#### Parameters

`int|array $rowNum`

---

### setRowOutlineLevel()

---

```php
public setRowOutlineLevel($rowNum, int $outlineLevel, ?bool $collapsed): Sheet
```


#### Parameters

`int|array|string $rowNum`

`int $outlineLevel`

`bool|null $collapsed`

---

### setRowStyle()

---

```php
public setRowStyle($rowRange, array $style): Sheet
```
_Style are applied to the entire sheet row (even if it is empty)_

#### Parameters

`int|string|array $rowRange`

`array $style`

---

### setRowStyleArray()

---

```php
public setRowStyleArray(array $rowStyles): Sheet
```
_Styles are applied to the entire sheet row (even if it is empty)_

#### Parameters

`array $rowStyles`

---

### setRowVisible()

---

```php
public setRowVisible($rowNum, bool $visible): Sheet
```
_Hide/show a specific row_

#### Parameters

`int|array $rowNum`

`bool $visible`

---

### skipRow()

---

```php
public skipRow(?int $rowCount): Sheet
```
_Skip rows_

#### Parameters

`int|null $rowCount`

---

### setStyle()

---

```php
public setStyle(string $cellAddress, $style, ?bool $mergeStyles): Sheet
```
_Alias for 'setCellStyle()'_

#### Parameters

`string $cellAddress`

`mixed $style`

`bool|null $mergeStyles`

---

### setTabColor()

---

```php
public setTabColor(?string $color): Sheet
```
_Set color for the sheet tab_

#### Parameters

`string|null $color`

---

### setTopLeftCell()

---

```php
public setTopLeftCell($cellAddress): Sheet
```
_Set top left cell for writing_

#### Parameters

`string|array $cellAddress`

---

### unprotect()

---

```php
public unprotect(): Sheet
```
_Unprotect sheet_

#### Parameters

_None_

---

### setValue()

---

```php
public setValue($cellAddress, $value, ?array $styles): Sheet
```
_Set a value to the single cell or to the cell range_

_$cellAddress formats:'B5''B5:C7'['col' => 2, 'row' => 5][2, 5]_

#### Parameters

`string|array $cellAddress`

`mixed $value`

`array|null $styles`

---

### withLastCell()

---

```php
public withLastCell(): Sheet
```
_Select last written cell for applying_

#### Parameters

_None_

---

### withLastRow()

---

```php
public withLastRow(): Sheet
```
_Select last written row for applying_

#### Parameters

_None_

---

### withRange()

---

```php
public withRange($range): Sheet
```
_Select custom range for applying_

#### Parameters

`array|string $range`

---

### writeAreas()

---

```php
public writeAreas(): Sheet
```


#### Parameters

_None_

---

### writeArray()

---

```php
public writeArray(array $rowArray = [], ?array $rowStyle): Sheet
```
_Write values from two-dimensional array_

#### Parameters

`array $rowArray` -- Array of rows

`array|null $rowStyle` -- Style applied to each row

---

### writeArrayTo()

---

```php
public writeArrayTo($topLeftCell, array $data): Sheet
```
_Write 2d array form the specified cell_

#### Parameters

`$topLeftCell`

`array $data`

---

### writeCell()

---

```php
public writeCell($value, ?array $styles): Sheet
```
_Write value to the current cell and move pointer to the next cell in the row_

#### Parameters

`mixed $value`

`array|null $styles`

---

### writeHeader()

---

```php
public writeHeader(array $header, ?array $rowStyle, ?array $colStyles = []): Sheet
```


#### Parameters

`array $header`

`array|null $rowStyle`

`array|null $colStyles`

---

### writeRow()

---

```php
public writeRow(array $rowValues = [], ?array $rowStyle, ?array $cellStyles): Sheet
```
_Write values to the current row_

#### Parameters

`array $rowValues` -- Values of cells

`array|null $rowStyle` -- Style applied to the entire row

`array|null $cellStyles` -- Styles of specified cells in the row

---

### writeTo()

---

```php
public writeTo($cellAddress, $value, ?array $styles = []): Sheet
```
_Write value to the specified cell and move pointer to the next cell in the row_

_$cellAddress formats:'B5''B5:C7'['col' => 2, 'row' => 5][2, 5]_

#### Parameters

`string|array $cellAddress`

`mixed $value`

`array|null $styles`

---

