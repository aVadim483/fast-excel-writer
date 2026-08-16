## FastExcelWriter – Workbook

### Workbook settings

```php
// Creates workbook with one default sheet 
$excel = Excel::create();

// Creates workbook with one sheet named 'Abc' 
$excel = Excel::create('Abc');

// Creates workbook with several named sheets 'Foo' and 'Bar'
$excel = Excel::create(['Foo', 'Bar']);

$font = [
    Style::FONT_NAME => 'Arial', 
    Style::FONT_SIZE => 14
];

// Creates workbook with default font style
$excel = Excel::create(['Foo', 'Bar'], [Style::FONT => $font]);

// Automatically convert strings containing numbers to numbers.
// Strings with a leading zero ('007') and numbers longer than 15 significant digits
// (IDs, barcodes) are kept as text - Excel would round them silently
$excel = Excel::create([], ['auto_convert_number' => true]);

// Saving strings to the shared string xml
$excel = Excel::create([], ['shared_string' => true]);
// or other way
$excel = Excel::create();
$excel->setSharedString();

// Sets locale
// In most cases, the locale is automatically set correctly,
// but sometimes you need to do it manually
$excel = Excel::create([], ['locale' => 'fr']);
// or other way
$excel = Excel::create();
$excel->setLocale('fr');

// Sets default formats of the workbook - they override formats of the locale
// and are used by the '@date', '@time' and '@datetime' styles
$excel = Excel::create([], [
    'default_date_format' => 'DD.MM.YYYY',
    'default_time_format' => 'HH:MM',
    'default_datetime_format' => 'DD.MM.YYYY HH:MM',
]);
// or other way
$excel = Excel::create();
$excel->setDefaultDateFormat('DD.MM.YYYY');
$excel->setDefaultTimeFormat('HH:MM');
$excel->setDefaultDateTimeFormat('DD.MM.YYYY HH:MM');

// Sets default font
$excel->setDefaultFont($font);

// Sets default styles
$excel->setDefaultStyle([Style::FONT => $font]);

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

### The Options Class

Instead of an array of options, you can pass an instance of the ```Options``` class
to ```Excel::create()``` — it has a fluent interface

```php
use \avadim\FastExcelWriter\Excel;
use \avadim\FastExcelWriter\Options;

$options = Options::create()
    ->tempDir('/path/to/temp/dir') // directory for temporary files
    ->tempPrefix('xlsx_') // custom prefix for temporary files
    ->autoConvertNumber() // automatically convert strings containing numbers to numbers
    ->sharedString() // save strings to the shared string xml
    ->locale('fr') // set locale
    ->defaultFont([Style::FONT_NAME => 'Arial', Style::FONT_SIZE => 14]) // set default font
    ->defaultDateFormat('DD.MM.YYYY') // format of the '@date' style
    ->defaultTimeFormat('HH:MM') // format of the '@time' style
    ->defaultDateTimeFormat('DD.MM.YYYY HH:MM') // format of the '@datetime' style
;

$excel = Excel::create(['Sheet1'], $options);
```

See also: [Options class](91-api-class-options.md)

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
### Shared Strings

By default, strings are written directly into the sheets (*inline strings*). With the ```'shared_string'```
option every string goes into a separate table (```sharedStrings.xml```) and the cells refer to it by index.

```php
$excel = Excel::create([], ['shared_string' => true]);
```

The trade-off is memory against the size of the uncompressed XML:

* inline strings are streamed, so the memory does not depend on the amount of data;
* shared strings need a table of all unique strings in memory until the file is saved
  (about 0.5 KB per unique string) - this is the only part of the writing that is not streamed.

Measured on 500 000 string cells (100 000 rows x 5 columns), PHP 8.4:

| data | option | xlsx | uncompressed XML | peak memory |
|---|---|---|---|---|
| 5 000 unique values, many repeats | inline | 2.39 MB | 49.5 MB | 4 MB |
| 5 000 unique values, many repeats | shared | 2.09 MB | 18.9 MB | 6 MB |
| all 500 000 values are unique | inline | 12.53 MB | 56.8 MB | 4 MB |
| all 500 000 values are unique | shared | 12.82 MB | 59.1 MB | 284 MB |

Note how little the size of the xlsx changes: the ZIP compression squeezes repeated strings
of the sheet anyway. Shared strings shrink the *uncompressed* XML (Excel opens such a file faster),
but on unique values they make even the XML bigger (an index in the cell plus the string in the table)
and cost a lot of memory.

See also: [Streaming mode and memory](03-writing.md#streaming-mode-and-memory)

**Rule of thumb:** keep the default for large exports. Turn shared strings on when the values come from
a limited dictionary (statuses, categories, names) and repeat a lot, or when the consumer of the file
requires them. The library cannot make this choice for you: in streaming mode the share of repeats
is unknown until all the data is written.

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

