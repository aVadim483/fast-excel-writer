## New
Sheet::isName($name): bool -- Case-insensitive name checking
Excel::setActiveSheet($name): Excel -- Set active (default) sheet by case-insensitive name

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

* Supports workbook and sheet protection with/without passwords (see [Protection of workbook and sheets](/docs/05-protection.md) )

New methods for cells
* Sheet::applyUnlock()
* Sheet::applyHide()
* Sheet::applyNamedRange(string $name)

New methods
* Excel::sheet() is alias of Excel::getSheet()
* Sheet::cell($cellAddress) - Select a single cell or to cell range in the current row
* Area::cell($cellAddress) - Select a single cell or to cell range in the current row
