# FastExcelWriter

Лёгкий, мощный и очень быстрый генератор XLSX-файлов на PHP, спроектированный для высокой скорости и минимального расхода памяти.

**FastExcelWriter** позволяет создавать файлы в XLSX-формате, совместимые с MS Excel (Office 2007+), LibreOffice, OpenOffice и др.,
поддерживает множество возможностей:

* Создание и запись множества листов в одном файле
* Поддерживает форматирование значений, сохранение формул и активных гиперссылок 
* Поддерживает стили для ячеек, строк и столбцов — цвета, границы, шрифты и т.д.
* Можно задавать высоту строк и ширину столбцов (включая авто‑подбор ширины)
* Можно добавлять заметки и изображения в XLSX‑файлы
* Поддерживает защиту книги и листов с паролем/без пароля
* Поддерживает параметры страницы — поля, формат бумаги для печати
* Создание и вставка нескольких диаграмм
* Поддержка валидации данных и условного форматирования

* [Быстрый старт](/docs/ru/index.md#быстрый-старт)
    * [Установка](/docs/ru/index.md#installation)
    * [Простой пример](/docs/ru/index.md#simple-example)
    * [Расширенный пример](/docs/ru/index.md#advanced-example)
    * [Вставка примечаний](/docs/ru/index.md#adding-notes)
    * [Вставка изображений](/docs/ru/index.md#adding-images)
* [Рабочая книга](/docs/ru/01-workbook.md)
    * [Рабочая книга settings](/docs/ru/01-workbook.md#workbook-settings)
    * [Sets metadata of workbook](/docs/ru/01-workbook.md#sets-metadata-of-workbook)
    * [Set Directory For Temporary Files](/docs/ru/01-workbook.md#set-directory-for-temporary-files)
    * [Helpers methods](/docs/ru/01-workbook.md#helpers-methods)
* [Листы](/docs/ru/02-sheets.md)
    * [Create, select and remove sheet](/docs/ru/02-sheets.md#create-select-and-remove-sheet)
    * [Sheet settings](/docs/ru/02-sheets.md#sheet-settings)
    * [Page settings](/docs/ru/02-sheets.md#page-settings)
    * [Row's settings](/docs/ru/02-sheets.md#rows-settings)
    * [Column's settings](/docs/ru/02-sheets.md#columns-settings)
    * [Automatic column widths](/docs/ru/02-sheets.md#automatic-column-widths)
    * [Group/outline rows and columns](/docs/ru/02-sheets.md#groupoutline-rows-and-columns)
    * [Define Named Ranges](/docs/ru/02-sheets.md#define-named-ranges)
    * [Freeze Panes and Autofilter](/docs/ru/02-sheets.md#freeze-panes-and-autofilter)
    * [Setting Active Sheet and Cells](/docs/ru/02-sheets.md#setting-active-sheet-and-cells)
    * [Print Settings](/docs/ru/02-sheets.md#print-settings)
* [Запись данных](/docs/ru/03-writing.md)
    * [Запись данных Row by Row vs Direct](/docs/ru/03-writing.md#writing-row-by-row-vs-direct)
    * [Direct Запись данных To Cells](/docs/ru/03-writing.md#direct-writing-to-cells)
    * [Запись данных Cell Values](/docs/ru/03-writing.md#writing-cell-values)
    * [Merging Cells](/docs/ru/03-writing.md#merging-cells)
    * [Cell Formats](/docs/ru/03-writing.md#cell-formats)
    * [Formulas](/docs/ru/03-writing.md#formulas)
    * [Hyperlinks](/docs/ru/03-writing.md#hyperlinks)
    * [Using Rich Text](/docs/ru/03-writing.md#using-rich-text)
* [Стили](/docs/ru/04-styles.md)
    * [Style Properties As Array](/docs/ru/04-styles.md#style-properties-as-array)
    * [Class Style for Defining Style Properties](/docs/ru/04-styles.md#class-style-for-defining-style-properties)
    * [Cell Стили](/docs/ru/04-styles.md#cell-styles)
    * [Row Стили](/docs/ru/04-styles.md#row-styles)
    * [Column Стили](/docs/ru/04-styles.md#column-styles)
    * [Other Columns Options](/docs/ru/04-styles.md#other-columns-options)
    * [Apply Стили (The Fluent Interface)](/docs/ru/04-styles.md#apply-styles-the-fluent-interface)
    * [Apply Borders](/docs/ru/04-styles.md#apply-borders)
    * [Apply Fonts](/docs/ru/04-styles.md#apply-fonts)
    * [Apply Colors](/docs/ru/04-styles.md#apply-colors)
    * [Apply Text Стили](/docs/ru/04-styles.md#apply-text-styles)
* [Диаграммы](/docs/ru/05-charts.md)
    * [Simple usage](/docs/ru/05-charts.md#simple-usage-of-chart)
    * [Combo charts](/docs/ru/05-charts.md#combo-charts)
    * [Multiple charts](/docs/ru/05-charts.md#multiple-charts)
    * [Chart types](/docs/ru/05-charts.md#chart-types)
    * [Useful Chart Methods](/docs/ru/05-charts.md#useful-chart-methods)
* [Защита of workbook and sheets](/docs/ru/06-protection.md)
    * [Рабочая книга protection](/docs/ru/06-protection.md#workbook-protection)
    * [Sheet protection](/docs/ru/06-protection.md#sheet-protection)
    * [Cells locking/unlocking](/docs/ru/06-protection.md#cells-lockingunlocking)
* [Data validation](/docs/ru/07-validation.md)
    * [Simple usage](/docs/ru/07-validation.md#simple-usage)
    * [Define filters](/docs/ru/07-validation.md#define-filters)
    * [Check type of value](/docs/ru/07-validation.md#check-type-of-value)
    * [Custom filters](/docs/ru/07-validation.md#custom-filters)
    * [All Data Проверка данных settings](/docs/ru/07-validation.md#all-data-validation-settings)
    * [More than 64K validation rules](/docs/ru/07-validation.md#more-than-64k-validation-rules)
* [Условное Formatting](/docs/ru/08-conditional.md)
    * [Simple usage](/docs/ru/08-conditional.md#simple-usage)
    * [General cell value check](/docs/ru/08-conditional.md#general-cell-value-check)
    * [Expressions](/docs/ru/08-conditional.md#expressions)
    * [Gradient fill depending on values (colorScale)](/docs/ru/08-conditional.md#gradient-fill-depending-on-values-colorscale)
    * [Data strip inside a cell (dataBar)](/docs/ru/08-conditional.md#data-strip-inside-a-cell-databar)
* [API Reference](/docs/ru/90-api-reference.md)
    * [Excel class](/docs/ru/91-api-class-excel.md)
    * [Sheet class](/docs/ru/92-api-class-sheet.md)
    * [RichText class](/docs/ru/93-api-class-rich-text.md)
    * [Chart class](94-api-class-chart.md)
    * [DataПроверка данных class](95-api-class-data-validation.md)
    * [Условное class](96-api-class-conditional.md)
    * [Style class](97-api-class-style.md)


## Быстрый старт

You can find usage examples below or in ```/demo``` folder

### Установка

Use `composer` to install **FastExcelWriter** into your project:

```
composer require avadim/fast-excel-writer
```


### Простой пример
```php
use \avadim\FastExcelWriter\Excel;

$data = [
    ['2003-12-31', 'James', '220'],
    ['2003-8-23', 'Mike', '153.5'],
    ['2003-06-01', 'John', '34.12'],
];

$excel = Excel::create(['Sheet1']);
$sheet = $excel->sheet();

// Записываем строку заголовков
$sheet->writeRow(['Date', 'Name', 'Amount']);

// Записываем данные
foreach($data as $rowData) {
    $rowOptions = [
        'height' => 20,
    ];
    $sheet->writeRow($rowData, $rowOptions);
}

$excel->save('simple.xlsx');
```
Также можно скачать сгенерированный файл компьютер пользователя (отправить файл в браузер)
```php
$excel = Excel::create();
$sheet = $excel->sheet();

$sheet->writeCell(12345); // пишем целое число
$sheet->writeCell(123.45); // пишем число с плавающей точкой
$sheet->writeCell('12345'); // пишем строку
$sheet->writeCell(true); // пишем булево значение
$sheet->writeCell(fn() => $sheet->getCurrentCell()); // в ячейку записывается результат выполнения функции

$excel->download('download.xlsx');
```

### Расширенный пример

```php
use \avadim\FastExcelWriter\Excel;

$head = ['Date', 'Name', 'Amount'];
$data = [
    ['2003-12-31', 'James', '220'],
    ['2003-8-23', 'Mike', '153.5'],
    ['2003-06-01', 'John', '34.12'],
];
$headStyle = [
    'font' => [
        'style' => 'bold'
    ],
    'text-align' => 'center',
    'vertical-align' => 'center',
    'border' => 'thin',
    'height' => 24,
];

$excel = Excel::create(['Sheet1']);
$sheet = $excel->sheet();

// Запись строки заголовков (задаем стили строки через массив)
$sheet->writeHeader($head, $headStyle);

// Тот же результат можно получить через "плавающий" интерфейс
// (получается более наглядно)
$sheet->writeHeader($head)
    ->applyFontStyleBold()
    ->applyTextAlign('center', 'center')
    ->applyBorder(Style::BORDER_STYLE_THIN)
    ->applyRowHeight(24);

// Задаем опции столбцов - формат и ширину (первый вариант)
$sheet
    ->setColFormats(['@date', '@text', '0.00'])
    ->setColWidths([12, 14, 5]);

// Второй вариант задания тех же опций столбцов
$sheet
    // используется буква столбца и опции
    ->setColDataStyle('A', ['format' => '@date', 'width' => 12])
    // букву столбца можно задать в любом регистре
    ->setColDataStyle('b', ['format' => '@text', 'width' => 24])
    // столбец модно задать в виде числа (порядковый номер)
    ->setColDataStyle(3, ['format' => '0.00', 'width' => 15, 'color' => '#090'])
;

// Третий вариант - опции можно многомерным массивом
$sheet
    ->setColDataStyle([
        'A' => ['format' => '@date', 'width' => 12],
        'B' => ['format' => '@text', 'width' => 24],
        'C' => ['format' => '0.00', 'width' => 15, 'color' => '#090'],
    ]);

$rowNum = 1;
foreach($data as $rowData) {
    $rowOptions = [
        'height' => 20,
    ];
    if ($rowNum % 2) {
        $rowOptions['fill-color'] = '#eee';
    }
    $sheet->writeRow($rowData, $rowOptions);
}

$excel->save('simple.xlsx');
```


###  Вставка заметок

Есть два вида комментариев в Excel - **Примечания** и **Заметки**
(см. [The difference between threaded comments and notes](https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3)).
Заметки (Notes) - это комментарии старого типа в Excel (обычно текст на желтых "листочках").
Вы можете добавить их к ячейкам с помощью метода ```addNote()```

```php

$sheet->writeCell('Text to A1');
$sheet->addNote('A1', 'This is a note for cell A1');

$sheet->writeCell('Text to B1')->addNote('This is a note for B1');
$sheet->writeTo('C4', 'Text to C4')->addNote('Note for C1');

// Если указать диапазон ячеек, то заметка будет добавлена в левую верхнюю ячейку
$sheet->addNote('E4:F8', "This note\nwill added to E4");

// Вы можете разбить текст на несколько строк символом \n
$sheet->addNote('D7', "Line 1\nLine 2");

```

Вы можете изменить некоторые параметры заметки. Допустимые параметры заметки:

* **width** — значение по умолчанию ```'96pt'```
* **height** — значение по умолчанию ```'55.5pt'```
* **fill_color** — значение по умолчанию ```'#FFFFE1'```
* **show** — значение по умолчанию ```false```

```php

$sheet->addNote('A1', 'This is a note for cell A1', 
    ['width' => '200pt', 'height' => '100pt', 'fill_color' => '#ffcccc']);

// Параметры "width" и "height" могут быть числовыми, по умолчанию эти значения указываются в пунктах.
// Значение параметра "fill_color" можно сократить до трех символов
$noteStyle = [
    'width' => 200, // эквивалент '200pt'
    'height' => 100, // эквивалент '100pt'
    'fill_color' => 'fcc', // эквивалент '#ffcccc'
];
$sheet->writeCell('Text to B1')->addNote('This is a note for B1', $noteStyle);

// Эта заметка отображается сразу при открытии листа
// (а не по наведению мыши, как это работает по умолчанию)
$sheet->addNote('C8', 'This note is always visible', ['show' => true]);
```

Также вы можете использовать форматированный текст в заметках

```php
$richText = new \avadim\FastExcelWriter\RichText('here is <c=f00>red</c> and <c=00f>blue</c> text');
$sheet->addNote('C8', $richText);
```

Более подробную информацию об использовании форматированного текста можно найти здесь:
[Using Rich Text](/docs/ru/03-writing.md#using-rich-text)

###  Вставка изображений

Вы можете вставить изображение на лист из локального файла, URL или строки изображения в формате base64

```php
$sheet->addImage($cell, $imageFile, $imageStyle);

// Вставить изображение в ячейку A1 из локального файла
$sheet->addImage('A1', 'path/to/file');

// Вставить изображение в ячейку A1 из URL
$sheet->addImage('A1', 'https://site.com/image.jpg');

// Вставить изображение в ячейку A1 из строки base64
$sheet->addImage('A1', 'data:image/jpeg;base64,/9j/4AAQ...');

// Вставить изображение в ячейку B2 и задать размер 150 пикселей (высота изменится пропорционально)
$sheet->addImage('B2', 'path/to/file', ['width' => 150]);

// Задать высоту 150 пикселей (ширина будет изменена пропорционально)
$sheet->addImage('C3', 'path/to/file', ['height' => 150]);

// Задать ширину и высоту вставляемого изображения в пикселях
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150]);

// Добавить ссылку к изображению
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150, 'hyperlink' => 'https://www.google.com/']);
```

Доступные опции вставляемого изображения:

* 'width' — ширина
* 'height' — высота
* 'hyperlink' — ссылка
* 'x' — горизонтальное смещение в пикселях
* 'y' — вертикальное смещение в пикселях

**ВАЖНО:** В MS Excel значение "x" не может быть больше ширины столбца родительской ячейки, 
а значение "y" не может быть больше высоты строки.

## **FastExcelWriter** vs **PhpSpreadsheet**

**PhpSpreadsheet** отличная библиотека для чтения и записи документов с поддержкой множества разных форматов.
**FastExcelWriter** используется только для записи и только в XLSX-формате, но делает это очень быстро
и с минимальными затратами памяти, поэтому может использоваться для создания огромных файлов.

**FastExcelWriter**:

* быстрее в 7-9 раз
* потребляет памяти меньше в 8-10 раз
* поддерживает создание огромных таблиц в 100K+ строк и более

Бенчмарк PhpSpreadsheet (генерация без стилей)

| Rows x Cols | Time      | Memory     |
|-------------|-----------|------------|
| 1000 x 5    | 0.98 sec  | 2,048 Kb   |
| 1000 x 25   | 4.68 sec  | 14,336 Kb  |
| 5000 x 25   | 23.19 sec | 77,824 Kb  |
| 10000 x 50  | 105.8 sec | 256,000 Kb |

Бенчмарк FastExcelWriter (генерация без стилей)

| Rows x Cols | Time      | Memory   |
|-------------|-----------|----------|
| 1000 x 5    | 0.19 sec  | 2,048 Kb |
| 1000 x 25   | 1.36 sec  | 2,048 Kb |
| 5000 x 25   | 3.61 sec  | 2,048 Kb |
| 10000 x 50  | 13.02 sec | 2,048 Kb |

## Хотите поддержать FastExcelWriter?

Если вы находите этот пакет полезным, вы можете пожертвовать на чашку кофе (или бутылочку пива).

* USDT (TRC20) TSsUFvJehQBJCKeYgNNR1cpswY6JZnbZK7
* USDT (ERC20) 0x5244519D65035aF868a010C2f68a086F473FC82b
* ETH 0x5244519D65035aF868a010C2f68a086F473FC82b

Или просто поставьте звёзду на GitHub :)