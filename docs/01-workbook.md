## FastExcelWriter - Workbook

### Workbook settings

```php
// Creates workbook with one default sheet 
$excel = Excel::create();

// Creates workbook with one sheet named 'Abc' 
$excel = Excel::create('Abc');

// Creates workbook with several named sheets 'Foo' and 'Bar'
$excel = Excel::create(['Foo', 'Bar']);

// Sets locale
// In most cases, the locale is automatically set correctly,
// but sometimes you need to do it manually
$excel->setLocale('fr');

// Sets default font
$excel->setDefaultFont($font);

// Sets default styles
$excel->setDefaultStyle($font);

// Sets RTL settings
$excel->setRightToLeft(true);

// Sets default filename for saving
$excel->setFileName('/path/to/out/file.xlsx');

// Saves workbook to default file 
$excel->save();

// Saves workbook to specified file 
$excel->save($filename);

// Download generated file to client (send to browser)
$excel->download('name.xlsx');

```

### Sets metadata of workbook

```php
$excel->setMetadata($key, $value);

// Shortcut methods
$excel->setTitle($title);
$excel->setSubject($subject);
$excel->setAuthor($author);
$excel->setCompany($company);
$excel->setDescription($description);
$excel->setKeywords($keywords);

```

### Set Directory For Temporary Files

The library uses temporary files to generate the XLSX-file. If not specified, they are created in the system temporary directory
or in the current execution directory. But you can set the directory for temporary files.

```php
use \avadim\FastExcelWriter\Excel;

Excel::setTempDir('/path/to/temp/dir'); // use this call before Excel::create()
$excel = Excel::create();

// Or alternative variant
$excel = Excel::create('SheetName', ['temp_dir' => '/path/to/temp/dir']);

```

### Helpers methods

These are static helper methods that you can use in your applications

```php
// Convert a column letter to a number (ONE based)
$number = Excel::colNumber('C'); // => 3
$number = Excel::colNumber('BZ'); // => 78

// Convert letter to index (ZERO based)
$number = Excel::colIndex('C'); // => 2
$number = Excel::colIndex('BZ'); // => 77

// Reverse conversion - from number to letter (ONE based)
$letter = Excel::colLetter(3); // => 'C'
$letter = Excel::colLetter(78); // => 'BZ'

// Make address from row and column (ONE based)
$address = Excel::cellAddress(8, 12); // => 'L8'
$address = Excel::cellAddress(8, 12, true); // => '$L$8'
$address = Excel::cellAddress(8, 12, true, false); // => '$L8'
$address = Excel::cellAddress(8, 12, false, true); // => 'L$8'

```

Returns to [README.md](/README.md)
