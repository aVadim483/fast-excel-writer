## V.6.4

* Sheet::writeCells($values, $cellStyles) - Write several values into cells of one row
* Conditional Formatting (see [Conditional Formatting](/docs/08-conditional.md))
* Changed: Ignoring formulas when calculating automatic column width
* Fixed https://github.com/aVadim483/fast-excel-writer/issues/94

## V.6.2

* New methods Sheet::getCurrentRow(), Sheet::getCurrentCol(), Sheet::getCurrentCell()
* New methods Sheet::applyIndentLeft(), Sheet::applyIndentRight(), Sheet::applyIndentDistributed()
* New method Sheet::applyFillGradient()
* Method Sheet::addImage() has been updated -- Add an image from a local file, URL, or base64 image string
* Data can be passed to a cell as a function
* Fixed page margins
* Fixed https://github.com/aVadim483/fast-excel-writer/issues/93 (the behavior has been changed)
* Fixed https://github.com/aVadim483/fast-excel-writer/issues/98 (the behavior has been changed)

## V.6.1

* Sheet::setRowOptions(), Sheet::setColOptions(), Sheet::setRowStyles() and Sheet::setColStyles() are deprecated
* Sheet::setRowStyle($rowNum, $style) - set style for row (or row range)
* Sheet::setRowStyleArray($rowStyle) - set styles for rows (or row range)
* Sheet::setRowDataStyle($rowNum, $style)
* Sheet::setRowDataStyleArray($rowStyle)
* Sheet::setColStyle($colLetter, $style) - set style for column (or column range)
* Sheet::setColStyleArray($colStyles) - set style for column (or column range)
* Sheet::setColDataStyle($colLetter, $style)
* Sheet::setColDataStyleArray($colStyles)

## V.6.0

* Data validation (see [Data validation](/docs/07-validation.md))
* All methods of Style were extracted into class StyleManager (constants remained in the class Style)
* PHP extension 'intl' is desirable but not required
* Other refactoring

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
