# FastExcelWriter

Лёгкая, мощная и очень быстрая библиотека для создания XLSX-таблиц на чистом PHP. Библиотека спроектирована так,
чтобы работать максимально быстро и потреблять минимум памяти.

**FastExcelWriter** создаёт таблицы в формате XLSX, совместимые с MS Excel (Office 2007+), LibreOffice, OpenOffice и другими,
и поддерживает множество возможностей:

* Принимает данные в кодировке UTF-8
* Несколько листов в одной книге
* Поддерживает валютные, датовые и числовые форматы ячеек, формулы и активные гиперссылки
* Поддерживает большинство параметров стилей для ячеек, строк и колонок — цвета, границы, шрифты и т.д.
* Можно задавать высоту строк и ширину колонок (включая автоматический расчёт ширины)
* Можно добавлять формулы, заметки и изображения в XLSX-файлы
* Поддерживает защиту книги и листов с паролем и без него
* Поддерживает параметры страницы — поля, размер бумаги
* Вставка нескольких диаграмм
* Поддерживает проверку данных (data validation) и условное форматирование

Быстрый переход:

* [Быстрый старт](index.md#quick-start)
    * [Установка](index.md#installation)
    * [Простой пример](index.md#simple-example)
    * [Расширенный пример](index.md#advanced-example)
    * [Добавление заметок](index.md#adding-notes)
    * [Добавление изображений](index.md#adding-images)
* [Книга (Workbook)](01-workbook.md)
* [Листы (Sheets)](02-sheets.md)
* [Запись данных](03-writing.md)
* [Стили](04-styles.md)
* [Диаграммы](05-charts.md)
* [Защита книги и листов](06-protection.md)
* [Проверка данных](07-validation.md)
* [Условное форматирование](08-conditional.md)
* [Руководство по обновлению](09-upgrade.md)
* [Справочник API](90-api-reference.md)
    * [Класс Excel](91-api-class-excel.md)
    * [Класс Options](91-api-class-options.md)
    * [Класс Sheet](92-api-class-sheet.md)
    * [Класс RichText](93-api-class-rich-text.md)
    * [Класс RichTextFragment](93-api-class-rich-text-fragment.md)
    * [Класс Chart](94-api-class-chart.md)
    * [Класс DataValidation](95-api-class-data-validation.md)
    * [Класс Conditional](96-api-class-conditional.md)
    * [Класс Style](97-api-class-style.md)
    * [Класс ImageStyle](97-api-class-image-style.md)


## Быстрый старт { #quick-start }

Примеры использования вы найдёте ниже, а также в папке ```/demo```

### Установка { #installation }

Установите **FastExcelWriter** в свой проект с помощью `composer`:

```
composer require avadim/fast-excel-writer
```


### Простой пример { #simple-example }
```php
use \avadim\FastExcelWriter\Excel;

$data = [
    ['2003-12-31', 'James', '220'],
    ['2003-8-23', 'Mike', '153.5'],
    ['2003-06-01', 'John', '34.12'],
];

$excel = Excel::create(['Sheet1']);
$sheet = $excel->sheet();

// Записываем заголовки
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
Также можно отдать сгенерированный файл клиенту (отправить в браузер)
```php
$excel = Excel::create();
$sheet = $excel->sheet();

$sheet->writeCell(12345); // запись целого числа
$sheet->writeCell(123.45); // запись числа с плавающей точкой
$sheet->writeCell('12345'); // запись строки
$sheet->writeCell(true); // запись логического значения
$sheet->writeCell(fn() => $sheet->getCurrentCell()); // запись результата функции

$excel->download('download.xlsx');
```

### Расширенный пример { #advanced-example }

```php
use \avadim\FastExcelWriter\Excel;
use \avadim\FastExcelWriter\Style\Style;

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

// Записываем строку заголовков (стиль задаётся массивом)
$sheet->writeHeader($head, $headStyle);

// Тот же результат с помощью fluent-интерфейса
$sheet->writeHeader($head)
    ->applyFontStyleBold()
    ->applyTextAlign('center', 'center')
    ->applyBorder(Style::BORDER_STYLE_THIN)
    ->applyRowHeight(24);

// Задаём параметры колонок — формат и ширину (первый способ)
$sheet
    ->setColFormats(['@date', '@text', '0.00'])
    ->setColWidths([12, 14, 5]);

// Второй способ задать параметры колонок
$sheet
    // колонка и параметры
    ->setColDataStyle('A', ['format' => '@date', 'width' => 12])
    // букву колонки можно писать в нижнем регистре
    ->setColDataStyle('b', ['format' => '@text', 'width' => 24])
    // колонку можно указать номером
    ->setColDataStyle(3, ['format' => '0.00', 'width' => 15, 'color' => '#090'])
;

// Третий способ — все параметры в многоуровневом массиве (ключи первого уровня указывают на колонки)
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


###  Добавление заметок { #adding-notes }

Сейчас в Excel есть два типа комментариев — **примечания** (comments) и **заметки** (notes)
(см. [The difference between threaded comments and notes](https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3)).
Заметки — это комментарии «старого стиля» (текст на светло-жёлтом фоне).
Вы можете добавлять заметки к любым ячейкам с помощью метода ```addNote()```

```php

$sheet->writeCell('Text to A1');
$sheet->addNote('A1', 'This is a note for cell A1');

$sheet->writeCell('Text to B1')->addNote('This is a note for B1');
$sheet->writeTo('C4', 'Text to C4')->addNote('Note for C1');

// Если указать диапазон ячеек, заметка будет добавлена к левой верхней ячейке
$sheet->addNote('E4:F8', "This note\nwill added to E4");

// Текст можно разбить на несколько строк
$sheet->addNote('D7', "Line 1\nLine 2");

```

У заметок можно менять некоторые параметры. Доступные параметры:

* **width** — значение по умолчанию ```'96pt'```
* **height** — значение по умолчанию ```'55.5pt'```
* **fill_color** — значение по умолчанию ```'#FFFFE1'```
* **show** — значение по умолчанию ```false```

```php

$sheet->addNote('A1', 'This is a note for cell A1', 
    ['width' => '200pt', 'height' => '100pt', 'fill_color' => '#ffcccc']);

// Параметры "width" и "height" могут быть числами, по умолчанию значения задаются в пунктах (pt)
// Параметр "fill_color" можно сокращать
$noteStyle = [
    'width' => 200, // эквивалентно '200pt'
    'height' => 100, // эквивалентно '100pt'
    'fill_color' => 'fcc', // эквивалентно '#ffcccc'
];
$sheet->writeCell('Text to B1')->addNote('This is a note for B1', $noteStyle);

// Эта заметка видна сразу при открытии книги
$sheet->addNote('C8', 'This note is always visible', ['show' => true]);
```

Также в заметках можно использовать форматированный текст (rich text)

```php
$richText = new \avadim\FastExcelWriter\RichText\RichText('here is <c=f00>red</c> and <c=00f>blue</c> text');
$sheet->addNote('C8', $richText);
```

Подробнее об использовании форматированного текста: [Форматированный текст](03-writing.md#using-rich-text)

###  Добавление изображений { #adding-images }

Вы можете вставить на лист изображение из локального файла, по URL или из строки в base64

```php
$sheet->addImage($cell, $imageFile, $imageStyle);

// Вставляем изображение в ячейку A1 из локального файла
$sheet->addImage('A1', 'path/to/file');

// Вставляем изображение в ячейку A1 по URL
$sheet->addImage('A1', 'https://site.com/image.jpg');

// Вставляем изображение в ячейку A1 из строки base64
$sheet->addImage('A1', 'data:image/jpeg;base64,/9j/4AAQ...');

// Вставляем изображение в ячейку B2 и задаём ширину 150 пикселей (высота изменится пропорционально)
$sheet->addImage('B2', 'path/to/file', ['width' => 150]);

// Задаём высоту 150 пикселей (ширина изменится пропорционально)
$sheet->addImage('C3', 'path/to/file', ['height' => 150]);

// Задаём размер в пикселях
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150]);

// Добавляем гиперссылку к изображению
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150, 'hyperlink' => 'https://www.google.com/']);
```

Доступные ключи стиля изображения:

* 'width' — ширина изображения
* 'height' — высота изображения
* 'hyperlink' — URL гиперссылки
* 'x' — смещение в пикселях относительно левой границы ячейки
* 'y' — смещение в пикселях относительно верхней границы ячейки

**ВАЖНО:** в MS Excel значение 'x' не может превышать ширину колонки родительской ячейки,
а значение 'y' — высоту строки

Вместо массива можно передать экземпляр класса ```ImageStyle``` — у него есть fluent-интерфейс

```php
use \avadim\FastExcelWriter\Style\ImageStyle;

$imageStyle = (new ImageStyle())
    ->width(150)
    ->height(150)
    ->offset(2, 3)
    ->hyperlink('https://www.google.com/');
$sheet->addImage('B2', 'path/to/file', $imageStyle);

// Или можно передать массив параметров в конструктор
$imageStyle = new ImageStyle(['width' => 150, 'height' => 150]);
$sheet->addImage('C3', 'path/to/file', $imageStyle);
```

См. также: [Класс ImageStyle](97-api-class-image-style.md)

## **FastExcelWriter** vs **PhpSpreadsheet**

**PhpSpreadsheet** — отличная библиотека с богатыми возможностями чтения и записи множества форматов документов.
**FastExcelWriter** умеет только записывать и только в формате XLSX, но делает это очень быстро
и с минимальным потреблением памяти.

**FastExcelWriter**:

* в 7–9 раз быстрее
* потребляет в 8–10 раз меньше памяти
* поддерживает запись огромных таблиц в 100K+ строк

Бенчмарк PhpSpreadsheet (генерация без стилей)

| Строк x Колонок | Время     | Память     |
|-----------------|-----------|------------|
| 1000 x 5        | 0.98 sec  | 2,048 Kb   |
| 1000 x 25       | 4.68 sec  | 14,336 Kb  |
| 5000 x 25       | 23.19 sec | 77,824 Kb  |
| 10000 x 50      | 105.8 sec | 256,000 Kb |

Бенчмарк FastExcelWriter (генерация без стилей)

| Строк x Колонок | Время     | Память   |
|-----------------|-----------|----------|
| 1000 x 5        | 0.19 sec  | 2,048 Kb |
| 1000 x 25       | 1.36 sec  | 2,048 Kb |
| 5000 x 25       | 3.61 sec  | 2,048 Kb |
| 10000 x 50      | 13.02 sec | 2,048 Kb |

## Хотите поддержать FastExcelWriter?

Если библиотека оказалась вам полезной, вы можете поддержать автора и задонатить ему на чашку кофе:

* USDT (TRC20) TSsUFvJehQBJCKeYgNNR1cpswY6JZnbZK7
* USDT (ERC20) 0x5244519D65035aF868a010C2f68a086F473FC82b
* ETH 0x5244519D65035aF868a010C2f68a086F473FC82b

Или просто поставьте звёздочку на GitHub :)
