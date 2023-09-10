## v.4.5

* Supports workbook and sheet protection with/without passwords (see [Protection of workbook and sheets](/docs/05-protection.md) )

New methods for cells
* Sheet::applyUnlock()
* Sheet::applyHide()
* Sheet::applyNamedRange()

New methods
* Excel::sheet() is alias of Excel::getSheet()
* Sheet::cell() - Select a single cell or to cell range in the current row
