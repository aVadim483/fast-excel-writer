## V.6.15

* New: a `Style` object can be created directly from an options array – `new Style($options)`
* New method `Excel::getSharedStringsRefCount()`
* Fixed reused `DataValidation`/`Conditional` objects – applying one object to several ranges kept only the last range, see https://github.com/aVadim483/fast-excel-writer/issues/137
* Fixed `Sheet::setStateVeryHidden()` – it made the sheet visible instead of very hidden
* Fixed `Sheet::writeCell()` – the last valid row (1048576) and column (XFD) were rejected
* Fixed `Excel::rangeDimensionRelative()` – a range given as two associative points collapsed to a single cell
* Fixed loss of a cell value containing invalid UTF-8 (`Writer::xmlSpecialChars()` now uses `ENT_SUBSTITUTE`)
* Fixed number/date output under a comma `LC_NUMERIC` locale (e.g. `de_DE`) on PHP 7.4 – no more invalid `<v>3,14</v>`
* Fixed style corruption when two workbooks are open in the same process (`StyleManager` read default-font state from the wrong instance)
* Fixed duplicate sheet names, and reuse of `sheetN.xml` after `removeSheet()` + `makeSheet()`
* Fixed missing XML escaping of image names, chart names and hyperlink targets/locations
* Fixed `Sheet::writeCells()` dropping positional (integer-keyed) cell styles
* Fixed formulas with escaped quotes `""` and national function names containing regex metacharacters
* Fixed shared strings losing leading/trailing spaces (`xml:space="preserve"`)
* Fixed `Area::getOffsetAddress()` ignoring its offset argument
* Fixed `Sheet::skipRow(null)` skipping no rows
* Fixed PHP 8.1+ deprecations from `strtolower()`/`htmlspecialchars()` on non-string arguments
* Fixed `ZipArchive` handle not being closed if saving throws an exception
* Security: `Excel::download()` strips CR/LF and quotes from the file name (header injection), sends the correct xlsx MIME type, adds an RFC 5987 `filename*` and checks `headers_sent()`
* Security: `Excel::loadImageFile()` rejects dangerous stream wrappers (`php://`, `phar://`, `glob://`, `zip://`, `data://`, `expect://`, …)
* Optimized formula conversion – the heavy `_xlfn`/`_xlws` passes are skipped when the formula contains no `(`
* Optimized `sharedStrings.xml` writing – streamed to disk instead of built entirely in memory (much lower peak memory when `shared_string` is enabled)
* Optimized the `Sheet::writeRow()` hot path
* Changed: after `removeSheet()` + `makeSheet()`, new worksheet files use a monotonic index and never reuse a previous `sheetN.xml`

## V.6.14

* New demos: rich text (demo-14), images & notes (demo-15); errors fixed in other demos, deprecated chart methods replaced
* Documentation: Russian translation, upgrade guide, ImageStyle and Options pages, rewritten README
* Optimized `Writer::_writeCell()`
* Fixed `Charts\Layout::getXMode()` and `Charts\Layout::getYMode()` – fatal TypeError when adding pie/donut charts
* Fixed `RichTextFragment::outXml()` – font size set via `setSize()` was not written

## V.6.13

* New class `ImageStyle`, can be used as an argument in `Sheet::addImage()`
* Fixed data validation formulas
* Fixed `Sheet::setColFormats()`, `Sheet::setColStyleArray()`, `Sheet::setColDataStyleArray()`
* V.6.13.1: fixed `Style::setVerticalAlign()`
* V.6.13.2: fixed https://github.com/aVadim483/fast-excel-writer/issues/136

## V.6.12

* New class `Options` for workbook settings
* New methods `Excel::getStyleCellStyles()` and `Excel::getStyleCellStyleXfs()`

## V.6.9

* IMPORTANT: The namespace of the RichText class has been changed to avadim\FastExcelWriter\RichText
* IMPORTANT: The namespace of the Style, StyleManager, and Font has been changed to avadim\FastExcelWriter\Style
* Improved hyperlink handling
* Deprecated methods removed: Sheet::setColStyles(), Sheet::setColOptions(), Sheet::getExternalLinks(), Sheet::setPageOptions(), Sheet::setRowOptions(), Sheet::setRowStyles(),
* An instance of the Style object can be used as an argument in all methods where an array with styles was passed

## V.6.8

* Sheet::setAutoFilter() – Supports specifying range
* Sheet::writeRows() – Write several rows from a two-dimensional array
* Sheet::nextRow() – fixed exporting of empty rows
* Fixed errors in php 8.5

## V.6.4

* Sheet::writeCells($values, $cellStyles) – Write several values into cells of one row
* Conditional Formatting (see [Conditional Formatting](/docs/08-conditional.md))
* Changed: Ignoring formulas when calculating automatic column width
* Fixed https://github.com/aVadim483/fast-excel-writer/issues/94

## V.6.2

* New methods Sheet::getCurrentRow(), Sheet::getCurrentCol(), Sheet::getCurrentCell()
* New methods Sheet::applyIndentLeft(), Sheet::applyIndentRight(), Sheet::applyIndentDistributed()
* New method Sheet::applyFillGradient()
* Method Sheet::addImage() has been updated – Add an image from a local file, URL, or base64 image string
* Data can be passed to a cell as a function
* Fixed page margins
* Fixed https://github.com/aVadim483/fast-excel-writer/issues/93 (the behavior has been changed)
* Fixed https://github.com/aVadim483/fast-excel-writer/issues/98 (the behavior has been changed)

## V.6.1

* Sheet::setRowOptions(), Sheet::setColOptions(), Sheet::setRowStyles() and Sheet::setColStyles() are deprecated
* Sheet::setRowStyle($rowNum, $style) – set style for row (or row range)
* Sheet::setRowStyleArray($rowStyle) – set styles for rows (or row range)
* Sheet::setRowDataStyle($rowNum, $style)
* Sheet::setRowDataStyleArray($rowStyle)
* Sheet::setColStyle($colLetter, $style) – set style for column (or column range)
* Sheet::setColStyleArray($colStyles) – set style for column (or column range)
* Sheet::setColDataStyle($colLetter, $style)
* Sheet::setColDataStyleArray($colStyles)

## V.6.0

* Data validation (see [Data validation](/docs/07-validation.md))
* All methods of Style were extracted into class StyleManager (constants remained in the class Style)
* PHP extension 'intl' is desirable but not required
* Another refactoring

## v.5.8

* $sheet->setTabColor($color);
* New Excel option - 'auto_convert_number';
* New Excel option - 'shared_string';
* New Excel option - 'locale';

## v.5.7

* $sheet->addImage($cell, $path, \['hyperlink' => $url\]);
* cell merge optimization

## v.5.6

* Excel::setActiveSheet($name): Excel -- Set active (default) sheet by case-insensitive name
* Sheet::isName($name): bool -- Case-insensitive name checking
* Sheet::setPrintArea($range): Sheet
* Sheet::setPrintTopRows($rows): Sheet
* Sheet::setPrintLeftColumns($cols): Sheet
* Sheet::setPrintGridlines($bool): Sheet

## v.5.5

* Support rich text in cells and notes
* Group/outline columns and rows

## v.5.3

* Sheet::writeArray()
* New option 'temp_prefix' - custom prefix for temporary files
* upd: Improved column width calculation
* fix: Note shape
* fix: R1C1 in formulas

## v.5.1

* Combo charts
* Custom colors
* Custom chart lines width
* Some methods have been renamed
  * setDataSeriesTickLabels() => setCategoryAxisLabels()
  * setXAxisLabel() => setCategoryAxisTitle()
  * getXAxisLabel() => getCategoryAxisTitle()
  * setYAxisLabel() => setValueAxisTitle()
  * getYAxisLabel() => getValueAxisTitle()

## v.5.0

Chart support!

## v.4.7

The order of writing entries to the file has been changed so that the mimetype is determined correctly

Bug fixes and some improvements

New method
* Sheet::setActiveCell($cellAddress);

## v.4.6

New methods
* Sheet::pageMargins(array $margins)
* Sheet::pageMarginLeft($value)
* Sheet::pageMarginRight($value)
* Sheet::pageMarginTop($value)
* Sheet::pageMarginBottom($value)
* Sheet::pageMarginHeader($value)
* Sheet::pageMarginFooter($value)

* Sheet::pagePaperSize($value)
* Sheet::pagePaperHeight($value)
* Sheet::pagePaperWidth($value)

* Sheet::setColAutoWidth($col) - alias of Sheet::setColWidthAuto($col)
* Sheet::setColMinWidth($col, $width)
* Sheet::setColMinWidths(array $widths)
* Sheet::setColVisible($col, bool $val)
* Sheet::setColHidden($col)

* Sheet::setRowVisible($rowNum, bool $val)
* Sheet::setRowHidden($rowNum)

Deprecated
* Sheet::setPageOptions()

## v.4.5

* Supports workbook and sheet protection with/without passwords (see [Protection of workbook and sheets](/docs/06-protection.md) )

New methods for cells
* Sheet::applyUnlock()
* Sheet::applyHide()
* Sheet::applyNamedRange(string $name)

New methods
* Excel::sheet() is alias of Excel::getSheet()
* Sheet::cell($cellAddress) - Select a single cell or to cell range in the current row
* Area::cell($cellAddress) - Select a single cell or to cell range in the current row
