# FastExcelWriter

## Border Styles

```php
// Create Excel workbook
$excel = Excel::create();

// Get the first sheet;
$sheet = $excel->getSheet();

// simple border style
$style1 = [
    Style::BORDER => Style::BORDER_THIN
];
$sheet->setStyle('A1', $style1);


```
