# FastExcelWriter

## Apply Styles

Methods that start with ```'apply...'``` are applied to either the cell or group of cells where the data was written.

```php
// Create Excel workbook
$excel = Excel::create();

// Get the first sheet;
$sheet = $excel->getSheet();

// The background color will be applied to all changed cells in the row
$sheet->writeRow(['foo', 'bar'])->applyBgColor('#f99');

// The background color will only be applied to the last changed cell
$sheet->writeCell('abc')->applyBgColor('#9f9');

// The background color will only be applied to cell C3
$sheet->writeTo('C3', 'edf')->applyBgColor('#cc99ff');

// Select the specified range and apply outer and inner cell borders for it
$sheet->withRange('B4:D5')->applyBgColor('#cff')->applyBorderOuter(Style::BORDER_DOUBLE)->applyBorderInner(Style::BORDER_DOTTED);

```

applyBorder(string $style, ?string $color = '#000000')
applyBorderLeft(string $style, ?string $color = '#000000')
applyBorderRight(string $style, ?string $color = '#000000')
applyBorderTop(string $style, ?string $color = '#000000')
applyBorderBottom(string $style, ?string $color = '#000000')
applyBorderOuter(string $style, ?string $color = '#000000')
applyBorderInner(string $style, ?string $color = '#000000')

applyFont(string $fontName, ?int $fontSize = null, ?string $fontStyle = null, ?string $fontColor = null)
applyFontName(string $fontName)
applyFontSize(float $fontSize)
applyFontStyle(string $fontStyle)
applyFontColor(string $fontColor)
applyFontStyleBold()
applyFontStyleItalic()
applyFontStyleStrikethrough()
applyFontStyleUnderline(?bool $double = false)

applyColor(string $color)
applyTextColor(string $color)
applyFillColor(string $color)
applyBgColor(string $color)

applyTextAlign(string $textAlign, ?string $verticalAlign = null)
applyVerticalAlign(string $verticalAlign)
applyTextCenter()
applyTextWrap(bool $textWrap)

## Columns options

```php
// Set widths of columns from the first (A)
$sheet->setColWidths([5, 16, 16, 'auto']);

// Set width of the column
$sheet->setColWidth(['G', 'H', 'J'], 14);

// Set formats of columns from the first (A); null - default format
$sheet->setColFormats([null, '@', '@', '@date', '0', '0.00', '@money', '@money']);

// Set style and width for specified column
$sheet->setColOptions('K', ['text-wrap' => true, 'width' => 32]);
```
