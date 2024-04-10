## FastExcelWriter - Protection

There are three levels of data protection in Excel spreadsheets:
* **Workbook** - You can protect a workbook from changing its structure by prohibiting insertion, deletion, renaming of sheets 
* **Sheet** - You can protect a sheet from modification, but at the same time allow certain actions with it, for example, selecting and sorting cells, changing the format or inserting rows, etc.
* **Cell** - If the worksheet is protected, then by default all the cells in the worksheet are protected from changes, but you can allow editing of certain cells, and you can also prevent cell formulas from being displayed in the protected worksheet

You can lock protection the workbook or the sheet by password. If a password is specified, 
the user can unprotect the workbook or the sheet only by knowing this password.

Please note that using a password does not encrypt the file and it can be read by any program. 
The password only blocks user actions to remove protection in the MS Excel interface. 
However, some compatible programs may ignore the password

### Workbook protection

```php
$excel = Excel::create();

// Protect a workbook
$excel->protect();

// Lock protection of the sheet by password
$excel->protect('PasswordForWorkbook');

// Unprotect a workbook
$excel->unprotect();

```

### Sheet protection

```php
$excel = Excel::create();

// Get the first sheet and protect one without password
$sheet1 = $excel->sheet();
$sheet1->protect();

// Lock protection of the sheet by password
$sheet2 = $excel->sheet('Sheet2');
$sheet2->protect('PasswordIsHere');

// Allow row insertion
$sheet2->allowInsertRows();

```
Methods that allow certain actions when the sheet is protected

| Function                   | Description                                                               |
|----------------------------|---------------------------------------------------------------------------|
| allowAutoFilter()          | AutoFilters should be allowed to operate when the sheet is protected      |
| allowDeleteColumns()       | Deleting columns should be allowed when the sheet is protected            |
| allowDeleteRows()          | Deleting rows should be allowed when the sheet is protected               |
| allowFormatCells()         | Formatting cells should be allowed when the sheet is protected            |
| allowFormatColumns()       | Formatting columns should be allowed when the sheet is protected          |
| allowFormatRows()          | Formatting rows should be allowed when the sheet is protected             |
| allowInsertColumns()       | Inserting columns should be allowed when the sheet is protected           |
| allowInsertHyperlinks()    | Inserting hyperlinks should be allowed when the sheet is protected        |
| allowInsertRows()          | Inserting rows should be allowed when the sheet is protected              |
| allowEditObjects()         | Objects are allowed to be edited when the sheet is protected              |
| allowPivotTables()         | PivotTables should be allowed to operate when the sheet is protected      |
| allowEditScenarios()       | Scenarios are allowed to be edited when the sheet is protected            |
| allowSelectLockedCells()   | Selection of locked cells should be allowed when the sheet is protected   |
| allowSelectUnlockedCells() | Selection of unlocked cells should be allowed when the sheet is protected |
| allowSelectCells()         | Selection of any cells should be allowed when the sheet is protected      |
| allowSort()                | Sorting should be allowed when the sheet is protected                     |

If the sheet is not protected, then these methods will be ignored

### Cells locking/unlocking

You can unlock particular cell in the protected sheet

```php
$sheet1->protect();

// Unlock cell when the sheet is protected
$sheet1->writeCell('')->applyBorder('thin')->applyUnlock();

// Hidden cell 
// Contents of the cell will not be displayed in the formula bar when the sheet is protected
// If the cell contains a formula then the cell should display the calculated result, but will not display the formula
$sheet1->writeCell('=SUM(A1:B3)')->applyHide();

```

You can unlock particular cell in the protected sheet

```php
$sheet1->protect();

// Unlock cell B4 when the sheet is protected
$sheet1->cell('B4')->applyBorder('thin')->applyUnlock();

// Hidden cell 
// Contents of the cell will not be displayed in the formula bar when the sheet is protected
// If the cell contains a formula then the cell should display the calculated result, but will not display the formula
$sheet1->writeCell('=SUM(A1:B3)')->applyHide();

```
Also, you can unlock cells for the any area

```php
$sheet1->protect();

$area1 = $sheet1->makeArea('C3:F9')
    ->applyOuterBorder('thin')
    ->applyBgColor('#ccc');
    
// Unlock cells inside an area    
$area2 = $sheet1->makeArea('d4:e8')
    ->applyOuterBorder('thin')
    ->applyBgColor('none')
    ->applyUnlock();

$sheet1->clearAreas();

```

