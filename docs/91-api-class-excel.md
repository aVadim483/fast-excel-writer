# Class \avadim\FastExcelWriter\Excel

---

* [__construct()](#__construct) -- Excel constructor
* [cellAddress()](#celladdress) -- Create cell address by row and col numbers
* [colIndex()](#colindex) -- Convert letter to index (ZERO based)
* [colIndexRange()](#colindexrange) -- Convert letter range to array of numbers (ZERO based)
* [colKeysToIndexes()](#colkeystoindexes)
* [colKeysToLetters()](#colkeystoletters)
* [colKeysToNumbers()](#colkeystonumbers)
* [colLetter()](#colletter) -- Convert column number to letter
* [colLetterRange()](#colletterrange) -- Convert values to letters array
* [colNumber()](#colnumber) -- Convert letter to number (ONE based)
* [colNumberRange()](#colnumberrange) -- Convert letter range to array of numbers (ONE based)
* [create()](#create) -- Create new workbook
* [createSheet()](#createsheet)
* [fullAddress()](#fulladdress)
* [generateUuid()](#generateuuid) -- Generate UUID v4
* [pixelsToEMU()](#pixelstoemu)
* [rangeRelOffsets()](#rangereloffsets) -- Return offsets by relative address (zero based)
* [rowIndexRange()](#rowindexrange)
* [rowNumberRange()](#rownumberrange)
* [setTempDir()](#settempdir) -- Set dir for temporary files
* [toTimestamp()](#totimestamp) -- Convert value (int or string) to Excel timestamp
* [setActiveSheet()](#setactivesheet) -- Set active (default) sheet by case-insensitive name
* [addDefinedName()](#adddefinedname)
* [addNamedRange()](#addnamedrange)
* [addSharedString()](#addsharedstring)
* [addStyle()](#addstyle)
* [addStyleDxfs()](#addstyledxfs)
* [setAuthor()](#setauthor)
* [setCompany()](#setcompany)
* [setDefaultFont()](#setdefaultfont) -- Set default font options
* [setDefaultFontName()](#setdefaultfontname) -- Set default font name
* [getDefaultFormatStyles()](#getdefaultformatstyles)
* [setDefaultLocale()](#setdefaultlocale) -- Set default locale from the current environment
* [getDefaultSheetName()](#getdefaultsheetname)
* [getDefaultStyle()](#getdefaultstyle)
* [setDefaultStyle()](#setdefaultstyle) -- Set default style
* [getDefinedNames()](#getdefinednames)
* [setDescription()](#setdescription)
* [download()](#download) -- Download generated file to client (send to browser)
* [getFileName()](#getfilename) -- Returns default filename
* [setFileName()](#setfilename) -- Sets default filename for saving
* [getHyperlinkStyle()](#gethyperlinkstyle)
* [getImageFiles()](#getimagefiles)
* [setKeywords()](#setkeywords)
* [loadImageFile()](#loadimagefile)
* [setLocale()](#setlocale) -- Set locale information
* [makeSheet()](#makesheet)
* [setMetaAuthor()](#setmetaauthor) -- Set metadata 'author'
* [setMetaCompany()](#setmetacompany) -- Set metadata 'company'
* [getMetadata()](#getmetadata) -- Get metadata
* [setMetadata()](#setmetadata) -- Set metadata
* [setMetaDescription()](#setmetadescription) -- Set metadata 'description'
* [setMetaKeywords()](#setmetakeywords) -- Set metadata 'keywords'
* [setMetaSubject()](#setmetasubject) -- Set metadata 'subject'
* [setMetaTitle()](#setmetatitle) -- Set metadata 'title'
* [output()](#output) -- Alias of download()
* [protect()](#protect) -- Protect workbook
* [removeSheet()](#removesheet) -- Removes sheet by index or name of sheet.
* [isRightToLeft()](#isrighttoleft)
* [setRightToLeft()](#setrighttoleft)
* [save()](#save) -- Save generated XLSX-file
* [getSharedStrings()](#getsharedstrings)
* [sheet()](#sheet) -- Returns sheet by number or name of sheet.
* [getSheet()](#getsheet) -- Alias of sheet()
* [getSheets()](#getsheets) -- Returns all sheets
* [getStyleDxfs()](#getstyledxfs)
* [setSubject()](#setsubject)
* [setTitle()](#settitle)
* [unprotect()](#unprotect) -- Unprotect workbook
* [getWriter()](#getwriter)

---

## __construct()

---

```php
public function __construct(?array $options = [])
```
_Excel constructor_

### Parameters

* `array|null $options` -- Optional parameters: \['temp_dir' => ..., 'temp_prefix' => ..., 'auto_convert_number' => ..., 'shared_string' => ...]

---

## cellAddress()

---

```php
public static function cellAddress(int $rowNumber, int $colNumber, 
                                   ?bool $absolute = false, 
                                   ?bool $absoluteRow = null): string
```
_Create cell address by row and col numbers_

### Parameters

* `int $rowNumber` -- ONE based
* `int $colNumber` -- ONE based
* `bool|null $absolute`
* `bool|null $absoluteRow`

---

### Examples

```php
cellAddress(3, 3) => 'C3'
cellAddress(43, 27) => 'AA43'
cellAddress(43, 27, true) => '$AA$43'
cellAddress(43, 27, false, true) => 'AA$43'
```


---

## colIndex()

---

```php
public static function colIndex(string $colLetter): int
```
_Convert letter to index (ZERO based)_

### Parameters

* `string $colLetter`

---

## colIndexRange()

---

```php
public static function colIndexRange($colLetter): array
```
_Convert letter range to array of numbers (ZERO based)_

### Parameters

* `string|int|array $colLetter` -- e.g.: 'B', 2, 'C:F', \['A', 'B', 'C']

---

## colKeysToIndexes()

---

```php
public static function colKeysToIndexes(array $data, $offset): array
```


### Parameters

* `array $data`
* `$offset`

---

## colKeysToLetters()

---

```php
public static function colKeysToLetters(array $data): array
```


### Parameters

* `array $data`

---

## colKeysToNumbers()

---

```php
public static function colKeysToNumbers(array $data): array
```


### Parameters

* `array $data`

---

## colLetter()

---

```php
public static function colLetter(int $colNumber): string
```
_Convert column number to letter_

### Parameters

* `int $colNumber` -- ONE based

---

## colLetterRange()

---

```php
public static function colLetterRange($colKeys, ?int $baseNum = 0): array
```
_Convert values to letters array_

### Parameters

* `array|string $colKeys`
* `int|null $baseNum` -- 0 or 1

---

### Examples

```php
$res = colLetterRange([0, 1, 2]);    // returns ['A', 'B', 'C']
$res = colLetterRange([1, 2, 3], 1); // returns ['A', 'B', 'C']
$res = colLetterRange('B, E, F');    // returns ['B', 'E', 'F']
$res = colLetterRange('B-E, F');     // returns ['B', 'C', 'D', 'E', 'F']
$res = colLetterRange('B1-E8');      // returns ['B', 'C', 'D', 'E']
$res = colLetterRange('B1:E8');      // returns ['B:E']
```


---

## colNumber()

---

```php
public static function colNumber(string $colLetter): int
```
_Convert letter to number (ONE based)_

### Parameters

* `string $colLetter`

---

## colNumberRange()

---

```php
public static function colNumberRange($colLetter): array
```
_Convert letter range to array of numbers (ONE based)_

### Parameters

* `string|int|array $colLetter` -- e.g.: 'B', 2, 'C:F', \['A', 'B', 'C']

---

## create()

---

```php
public static function create($sheets, ?array $options = []): Excel
```
_Create new workbook_

### Parameters

* `array|string|null $sheets` -- Name of sheet or array of names
* `array|null $options` -- Options

---

## createSheet()

---

```php
public static function createSheet(string $sheetName): Sheet
```


### Parameters

* `string $sheetName`

---

## fullAddress()

---

```php
public static function fullAddress(string $sheetName, string $address, 
                                   ?bool $force = false): string
```


### Parameters

* `string $sheetName`
* `string $address`
* `bool|null $force`

---

## generateUuid()

---

```php
public static function generateUuid(): string
```
_Generate UUID v4_

### Parameters

_None_

---

## pixelsToEMU()

---

```php
public static function pixelsToEMU($pixels): float|int
```


### Parameters

* `$pixels`

---

## rangeRelOffsets()

---

```php
public static function rangeRelOffsets(string $relAddress): array
```
_Return offsets by relative address (zero based)_

### Parameters

* `string $relAddress`

---

## rowIndexRange()

---

```php
public static function rowIndexRange($rowRange): array
```


### Parameters

* `int|string|array $rowRange`

---

## rowNumberRange()

---

```php
public static function rowNumberRange($rowRange): array
```


### Parameters

* `int|string|array $rowRange`

---

## setTempDir()

---

```php
public static function setTempDir($tempDir)
```
_Set dir for temporary files_

### Parameters

* `$tempDir`

---

## toTimestamp()

---

```php
public static function toTimestamp($value): float|bool
```
_Convert value (int or string) to Excel timestamp_

### Parameters

* `int|string $value`

---

## setActiveSheet()

---

```php
public function setActiveSheet(string $name): Excel
```
_Set active (default) sheet by case-insensitive name_

### Parameters

* `string $name`

---

## addDefinedName()

---

```php
public function addDefinedName(string $name, string $range, 
                               ?array $attributes = []): Excel
```


### Parameters

* `string $name`
* `string $range`
* `array|null $attributes`

---

## addNamedRange()

---

```php
public function addNamedRange(string $range, string $name): Excel
```


### Parameters

* `string $range`
* `string $name`

---

## addSharedString()

---

```php
public function addSharedString(string $string, ?bool $richText = false): int
```


### Parameters

* `string $string`
* `bool|null $richText`

---

## addStyle()

---

```php
public function addStyle($cellStyle, &$resultStyle): int
```


### Parameters

* `$cellStyle`
* `$resultStyle`

---

## addStyleDxfs()

---

```php
public function addStyleDxfs($style, &$resultStyle): int
```


### Parameters

* `$style`
* `$resultStyle`

---

## setAuthor()

---

```php
public function setAuthor(?string $author = ''): Excel
```


### Parameters

* `$author`

---

## setCompany()

---

```php
public function setCompany(?string $company = ''): Excel
```


### Parameters

* `$company`

---

## setDefaultFont()

---

```php
public function setDefaultFont(array $fontOptions): Excel
```
_Set default font options_

### Parameters

* `array $fontOptions`

---

## setDefaultFontName()

---

```php
public function setDefaultFontName(string $fontName): Excel
```
_Set default font name_

### Parameters

* `string $fontName`

---

## getDefaultFormatStyles()

---

```php
public function getDefaultFormatStyles(): array
```


### Parameters

_None_

---

## setDefaultLocale()

---

```php
public function setDefaultLocale()
```
_Set default locale from the current environment_

### Parameters

_None_

---

## getDefaultSheetName()

---

```php
public function getDefaultSheetName(): string
```


### Parameters

_None_

---

## getDefaultStyle()

---

```php
public function getDefaultStyle(): array
```


### Parameters

_None_

---

## setDefaultStyle()

---

```php
public function setDefaultStyle(array $style): Excel
```
_Set default style_

### Parameters

* `array $style`

---

## getDefinedNames()

---

```php
public function getDefinedNames(): array
```


### Parameters

_None_

---

## setDescription()

---

```php
public function setDescription(?string $description = ''): Excel
```


### Parameters

* `$description`

---

## download()

---

```php
public function download(?string $name = null)
```
_Download generated file to client (send to browser)_

### Parameters

* `string|null $name`

---

## getFileName()

---

```php
public function getFileName(): string
```
_Returns default filename_

### Parameters

_None_

---

## setFileName()

---

```php
public function setFileName(string $fileName): Excel
```
_Sets default filename for saving_

### Parameters

* `string $fileName`

---

## getHyperlinkStyle()

---

```php
public function getHyperlinkStyle(): array
```


### Parameters

_None_

---

## getImageFiles()

---

```php
public function getImageFiles(): array
```


### Parameters

_None_

---

## setKeywords()

---

```php
public function setKeywords($keywords): Excel
```


### Parameters

* `$keywords`

---

## loadImageFile()

---

```php
public function loadImageFile(string $imageFile): ?array
```


### Parameters

* `string $imageFile` -- URL, local path or image string in base64

---

## setLocale()

---

```php
public function setLocale(string $locale, ?string $dir = null): Excel
```
_Set locale information_

### Parameters

* `string $locale`
* `string|null $dir`

---

## makeSheet()

---

```php
public function makeSheet(?string $sheetName = null): Sheet
```


### Parameters

* `string|null $sheetName`

---

## setMetaAuthor()

---

```php
public function setMetaAuthor(?string $author = ''): Excel
```
_Set metadata 'author'_

### Parameters

* `string|null $author`

---

## setMetaCompany()

---

```php
public function setMetaCompany(?string $company = ''): Excel
```
_Set metadata 'company'_

### Parameters

* `string|null $company`

---

## getMetadata()

---

```php
public function getMetadata($key): mixed
```
_Get metadata_

### Parameters

* `null $key`

---

## setMetadata()

---

```php
public function setMetadata($key, $value): Excel
```
_Set metadata_

### Parameters

* `$key`
* `$value`

---

## setMetaDescription()

---

```php
public function setMetaDescription(?string $description = ''): Excel
```
_Set metadata 'description'_

### Parameters

* `string|null $description`

---

## setMetaKeywords()

---

```php
public function setMetaKeywords($keywords): Excel
```
_Set metadata 'keywords'_

### Parameters

* `mixed $keywords`

---

## setMetaSubject()

---

```php
public function setMetaSubject(?string $subject = ''): Excel
```
_Set metadata 'subject'_

### Parameters

* `string|null $subject`

---

## setMetaTitle()

---

```php
public function setMetaTitle(?string $title = ''): Excel
```
_Set metadata 'title'_

### Parameters

* `string|null $title`

---

## output()

---

```php
public function output(?string $name = null): void
```
_Alias of download()_

### Parameters

* `string|null $name`

---

## protect()

---

```php
public function protect(?string $password = null): Excel
```
_Protect workbook_

### Parameters

* `string|null $password`

---

## removeSheet()

---

```php
public function removeSheet($index): Excel
```
_Removes sheet by index or name of sheet.Removes the first sheet of index omitted_

### Parameters

* `int|string|null $index`

---

## isRightToLeft()

---

```php
public function isRightToLeft(): bool
```


### Parameters

_None_

---

## setRightToLeft()

---

```php
public function setRightToLeft(bool $isRightToLeft)
```


### Parameters

* `bool $isRightToLeft`

---

## save()

---

```php
public function save(?string $fileName = null, ?bool $overWrite = true): bool
```
_Save generated XLSX-file_

### Parameters

* `string|null $fileName`
* `bool|null $overWrite`

---

## getSharedStrings()

---

```php
public function getSharedStrings(): array
```


### Parameters

_None_

---

## sheet()

---

```php
public function sheet($index): ?Sheet
```
_Returns sheet by number or name of sheet.Return the first sheet if number or name omitted_

### Parameters

* `int|string|null $index` -- - number or name of sheet

---

## getSheet()

---

```php
public function getSheet($index): ?Sheet
```
_Alias of sheet()_

### Parameters

* `int|string|null $index` -- - number or name of sheet

---

## getSheets()

---

```php
public function getSheets(): array
```
_Returns all sheets_

### Parameters

_None_

---

## getStyleDxfs()

---

```php
public function getStyleDxfs(): array
```


### Parameters

_None_

---

## setSubject()

---

```php
public function setSubject(?string $subject = ''): Excel
```


### Parameters

* `$subject`

---

## setTitle()

---

```php
public function setTitle(?string $title = ''): Excel
```


### Parameters

* `$title`

---

## unprotect()

---

```php
public function unprotect(): Excel
```
_Unprotect workbook_

### Parameters

_None_

---

## getWriter()

---

```php
public function getWriter(): Writer
```


### Parameters

_None_

---

