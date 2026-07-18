## FastExcelWriter – Стили

Стили можно задавать для всей книги, для листов, для отдельных ячеек, а также для строк и колонок.
Итоговый стиль каждой ячейки определяется комбинацией всех этих стилей.

### Свойства стиля в виде массива

```php
$style = [
    'format' => '#,##0.00',
    'font-color' => '#ff0000',
    'text-align' => 'center',
];
$sheet->writeCell(0.9, $style);
```

### Класс Style для определения свойств стиля

```php
$style = Excel::newStyle()
    ->setFillGradient('#fff000', '#fff')
    ->setFormat('#,##0.00')
    ->setFontColor('#ff0000')
    ->setTextAlign(Style::TEXT_ALIGN_CENTER)
;
 
$sheet->writeCell(0.9, $style->toArray());
```

### Стили ячеек

```php
$style = [
    'format' => '#,##0.00',
    'font-color' => '#ff0000',
    'text-align' => 'center',
];
$sheet->writeCell(0.9, $style);
$sheet->writeTo('B4', $value, $style);

// Задаёт стиль указанной ячейке
$sheet->setStyle('C8', $style);

// Начинаем область прямой записи
$area = $sheet->beginArea();
$area
    ->setValue('C10', 1234, $otherStyle)
    ->setValue('E12:K12', 'This is string', $alterStyle);

```

### Стили строк

```php
$rowStyle = [
    Style::FILL_COLOR => '#ff99ff',
    Style::BORDER => [
        Style::BORDER_TOP => [
            Style::BORDER_STYLE => Style::BORDER_THICK,
            Style::BORDER_COLOR => '#f00',
        ]
    ]
];
// Записываем данные строки и задаём её стили
$sheet->writeRow($rowData, $rowStyle);

// Задаём стиль для строки 3
$sheet->setRowStyle(3, $style);

$rowStyles = [
    3 => ['fill-color' => '#cff'], // параметры строки 3 
    4 => ['fill-color' => '#ccc', 'height' => 20], // параметры строки 4
];

// Задаём стили указанным строкам 
$sheet->setRowStyleArray($rowStyles);

// Задаём параметры диапазону строк 
$sheet->setRowStyle('2:5', ['fill-color' => '#f00']);

```

Стиль можно задать как для всей строки листа, так и только для тех ячеек строки, в которые записаны данные.
```php
// Стили применяются ко всей строке листа
$sheet->setRowStyle(3, ['height' => 20]);
$sheet->setRowStyle('2:5', ['font-color' => '#f00']);
$sheet->setRowStyleArray([3 => $style1, 5 => $style2]);

// Задаём стиль только для непустых ячеек строки
$sheet->setRowDataStyle(3, ['height' => 20]);
$sheet->setRowDataStyle('2:5', ['font-color' => '#f00']);
$sheet->setRowDataStyleArray([3 => $style1, 5 => $style2]);
```


### Стили строк и ячеек

Первый аргумент функции ```writeRow()``` — массив значений, второй — стиль строки, третий — стили ячеек строки.

```php
$rowValues = [
    'text',
    'http://google.com',
    123456,
];
$rowStyle = [];
$cellStyles = [
    [], // стиль первой ячейки,
    ['hyperlink' => true], // второй ячейки
    [], // третьей ячейки
];

$sheet->writeRow($rowValues, $rowStyle, $cellStyles);

```

Также можно задать стили для указанных ячеек

```php
$cellStyles = [
    'B' => ['hyperlink' => true],
];
$sheet->writeRow($rowValues, [], $cellStyles);
```

### Стили колонок

Стиль и формат колонки можно определить методом ```writeHeader()```.
Код ниже делает следующее:

* Записывает в ячейки текущей строки значения 'column title 1', 'column title 2', 'column title 3'
* Задаёт этой строке полужирный шрифт и тонкие границы (цвет границ по умолчанию — #000000)
* Задаёт стили, ширину и форматы колонок 'A', 'B' и 'C'

```php
$headValues = [
    // <значение ячейки> => [<стиль колонки>]
    ['column title 1' => ['format' => '@text', 'width' => 20, 'fill-color' => '#ccf']],
    ['column title 2' => ['format' => '@integer', 'width' => 12, 'font-color' => '#009']],
    ['column title 3' => ['text-wrap' => true, 'fill-color' => '#ccf']],
];
$headStyles = [
    'font-style' => 'bold',
    'border-style' => 'thin',
];
$sheet->writeHeader($headValues, $headStyles);

```
Можно задать стили для указанных колонок

```php
$sheet->setColStyle('C', $colStyle);
$sheet->setColWidth('E', 32);
$sheet->setColFormat('K', '@date');

```

Стиль можно задать как для всей колонки листа, так и только для тех ячеек колонки, в которые записаны данные.
```php
// Стили применяются ко всей колонке листа
$sheet->setColStyle('B', $style); // стиль для ячеек колонки 'B'
$sheet->setColStyle(2, $style); // колонка 'B' имеет номер 2
$sheet->setColStyle('C:F', $style); // стиль для диапазона колонок
$sheet->setColStyle(['A', 'B', 'C'], $style); // параметры для нескольких колонок 'A', 'B' и 'C'
$sheet->setColStyleArray(['B' => ['width' => 20], 'C' => ['font-color' => '#f00']]);

// Задаём стиль только для непустых ячеек колонки
$sheet->setColDataStyle('B', ['width' => 20]);
$sheet->setColDataStyle(2, ['width' => 20]);
$sheet->setColDataStyle('B:D', ['width' => 'auto']);
$sheet->setColDataStyle(['A', 'B', 'C'], $style);
$sheet->setColDataStyleArray(['B' => $style1, 'C' => $style2]);
```

### Другие параметры колонок

```php
// Задаём стиль и ширину указанной колонки
$sheet->setColDataStyle('K', ['text-wrap' => true, 'width' => 32]);

// Задаём ширину колонок начиная с первой (A)
$sheet->setColWidths([5, 16, 16, 'auto']);

// Задаём ширину колонок
$sheet->setColWidth(['G', 'H', 'J'], 14);

// Задаём форматы колонок начиная с первой (A); null — формат по умолчанию
$sheet->setColFormats([null, '@', '@', '@date', '0', '0.00', '@money', '@money']);

```

### Применение стилей (fluent-интерфейс)

Методы, начинающиеся с ```'apply...'```, применяются к ячейке или группе ячеек, в которые были записаны данные.

```php
// Создаём книгу
$excel = Excel::create();

// Получаем первый лист;
$sheet = $excel->getSheet();

// Цвет фона будет применён ко всем изменённым ячейкам строки
$sheet->writeRow(['foo', 'bar'])->applyBgColor('#f99');

// Цвет фона будет применён только к последней изменённой ячейке
$sheet->writeCell('abc')->applyBgColor('#9f9');

// Цвет фона будет применён только к ячейке C3
$sheet->writeTo('C3', 'edf')->applyBgColor('#cc99ff');

// Выбираем указанный диапазон и применяем к нему внешние и внутренние границы ячеек
$sheet->withRange('B4:D5')->applyBgColor('#cff')->applyBorderOuter(Style::BORDER_DOUBLE)->applyBorderInner(Style::BORDER_DOTTED);

```

#### Применение границ
* applyBorder(string $style, ?string $color = '#000000')
* applyBorderLeft(string $style, ?string $color = '#000000')
* applyBorderRight(string $style, ?string $color = '#000000')
* applyBorderTop(string $style, ?string $color = '#000000')
* applyBorderBottom(string $style, ?string $color = '#000000')
* applyBorderOuter(string $style, ?string $color = '#000000')
* applyBorderInner(string $style, ?string $color = '#000000')

#### Применение шрифтов
* applyFont(string $fontName, ?int $fontSize = null, ?string $fontStyle = null, ?string $fontColor = null)
* applyFontName(string $fontName)
* applyFontSize(float $fontSize)
* applyFontStyle(string $fontStyle)
* applyFontColor(string $fontColor)
* applyFontStyleBold()
* applyFontStyleItalic()
* applyFontStyleStrikethrough()
* applyFontStyleUnderline(?bool $double = false)

#### Применение цветов
* applyColor(string $color)
* applyTextColor(string $color)
* applyFillColor(string $color)
* applyBgColor(string $color)

#### Применение стилей текста
* applyTextAlign(string $textAlign, ?string $verticalAlign = null)
* applyVerticalAlign(string $verticalAlign)
* applyTextCenter()
* applyTextWrap(bool $textWrap)
* applyTextRotation(int $degrees) (спасибо @jarrod-colluco)
