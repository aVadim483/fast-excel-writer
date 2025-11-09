# Создание и форматирование Excel-файлов в FastExcelWriter

Содержание
Создание документа
Размеры ячеек
Объединение ячеек
Стили текста
Выравнивание в ячейке
Фон ячейки
Границы
Изображения
Гиперссылки
Формулы
Сохранение

## Создание документа

```php
$excelWriter = Excel::create();
```

### Установка метаданных документа:
```php
$excelWriter->setTitle("Название");
$excelWriter->setSubject("Тема");
$excelWriter->setAuthor("Автор");
$excelWriter->setCompany("Организация");
$excelWriter->setDescription("Описание");
$excelWriter->setKeywords("Ключевые слова");
```

### Работа с листами книги

Можно создать необходимые листы при создании книги Excel.
```php
// создание книги сразу с тремя листами
$excelWriter = Excel::create(['Январь', 'Февраль', 'Март']);
// получение первого листа по умолчанию
$sheet = $excelWriter->sheet();
// получение листа по его имени
$sheet = $excelWriter->sheet('Февраль');
// создание нового листа
$sheet = $excel->makeSheet('Апрель');
```
Основная работа записи и форматировании данных выполняется через экземпляр листа (Sheet),
в нашем примере — через переменную $sheet.

Лист можно переименовать
```php
$sheet->setName('Апрель 2025');
```

Задаем параметры печати, если нужно
```php
// Формат печатного листа
$sheet->pagePaperSize(Excel::PAPERSIZE_A3);

// Ориентация
$sheet->pagePortrait(); // книжная
$sheet->pageLandscape(); // альбомная

// Поля в дюймах
$sheet->pageMarginTop(1);
$sheet->pageMarginRight(0.75);
$sheet->pageMarginLeft(0.75);
$sheet->pageMarginBottom(1);

// Можно задать поля в см или мм
$sheet->pageMarginTop('1cm');
$sheet->pageMarginRight('30mm');

// Верхний колонтитул
$sheet->pageHeader("Заголовок");

// Нижний колонтитул
$sheet->pageFooter(['Текст слева', 'Название листа по центру', 'Страница &P из &N']);

```
### Запись данных

По умолчанию значения пишутся в ячейки последовательно, друг за другом, строка за строкой 
(построчная запись, которая как раз и дает мощный прирост по скорости и экономию памяти)

```php
$sheet = $excelWriter->sheet();
$sheet->writeCell("Значение ячейки A1");
$sheet->writeCell("Значение ячейки B1");
$sheet->nextRow(); // переход на следующую строку

$sheet->writeCell("Значение ячейки A2");
$sheet->writeCell("Значение ячейки B2");
$sheet->nextCell(); // переход к следующей ячейке 
$sheet->writeCell("Значение ячейки D2");

```
Разумеется, можно писать сразу строками, не обязательно по отдельным ячейкам
```php
$sheet->writeRow(['aaa', 'bbb', 'ccc']);
```

Можно записать сразу несколько строк из двумерного массива
```php
$sheet->writeRows([['aaa', 'bbb', 'ccc'], ['ddd', 'eee', 'fff']]);
```
При необходимости вы можете указать адрес ячейка, куда нужно записать значение.
Но важно помнить, что номер строки целевой ячейки не должен быть меньше уже записанных строк.

```php
// записываем строки 1 и 2
$sheet->writeRow([['aaa', 'bbb', 'ccc'], ['ddd', 'eee', 'fff']]);
// записываем в ячейку F4
$sheet->writeTo('F4', '...');
// но если теперь сделаем попытку записать в строку 3, то получим ошибку
$sheet->writeTo('A3', '...');

```
Значения в ячейки записываются в том виде, в каком передаются

```php
$sheet->writeCell(123); // будет записано число
$sheet->writeCell('123'); // будет записана строка
```
Но вы можете явно задавать формат ячеек
```php
$sheet->writeCell(123)->applyFormat("# ##0.000"); // задаем формат для числа
$sheet->writeCell(123)->applyFormat('@money'); // денежный формат
$sheet->writeCell('2025-01-28')->applyFormat('@date'); // указываем, что это дата, а не просто строка

```
Есть несколько предопределенных форматов, которые можно использовать, они начинаются с символа '@'

| Формат    | Значение     |
|-----------|--------------|
| @text     | Текст        |
| @string   | Текст        |
| @integer  | Целое число  |
| @date     | Дата         |
| @datetime | Дата и время |
| @time     | Время        |
| @money    | Деньги       |

Формат ячеек можно задавать сразу для всего столбца:

```php
$sheet->setColFormat('A', '@date');
```
Теперь для всех значений, которые записываются в столбец 'A' будет задаваться формат даты.

Можно задавать формат сразу для нескольких столбцов
```php
$sheet->setColFormat(['A', 'D', 'EX'], '@date');
$sheet->setColFormats(['A' => '@date', 'D'=> '# ##0.000']);
```

## Размеры ячеек

### Ширина столбцов и высота строк
```php
// Ширина столбца A
$sheet->setColWidth('A', 100);
// То же самое
$sheet->setColWidth(1, 100);

// Одинаковая ширина для нескольких столбцов
$sheet->setColWidth(['A', 'C', 'E'], 100);

// Разная ширина для разных столбцов (для 'C' задается автоширина)
$sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);

```
Можно задать и автоширину столбца по содержимому. Но важно понимать, что ширина вычисляется приблизительно, 
а если ячейка содержит формулу, то и вовсе может быть подсчитана некорректно.
```php
// Автоширина столбца 'B'
$sheet->setColWidthAuto(2);
$sheet->setColWidthAuto('B');

// Автоширина столбцов 'B' и 'C'
$sheet->setColWidthAuto(['B', 'C']);

// Автоширина всех столбцов в диапазоне от 'B' до 'E'
$sheet->setColWidthAuto(['B:E']);

```
Высота строки задается аналогично ширине столбца
```php
// Высота для заданной строки
$sheet->setRowHeight(1, 20);

// Задаем высоту для нескольких строк
$sheet->setRowHeight([1, 3, 24], 20);

// Задаем высоту для диапазона строк
$sheet->setRowHeight('3:7', 30);
```

## Объединение ячеек

Можно сначала записать все необходимые данные, а потом указать, какие ячейки нужно объединить
```php
$sheet->mergeCells('A6:C7');
```
А можно задавать объединение ячеек сразу при записи
```php
$sheet->writeTo('G3:I4', 'Hello');
```
ВАЖНО: формат XLSX не допускает пересечения адресов объединенных ячеек, при открытии такого файла Excel выдаст ошибку.
Чтобы это предотвратить, при вызове методов ```mergeCells()``` и ```writeTo()``` с объединением ячеек каждый раз проверяется,
нет ли пересечений с уже объединенными ячейками. И каждый раз выполняется проверка по полному списку объединенных ячеек.

Например, если вы записываете 1000 строк и в каждой строке у вас есть объединенные ячейки, то по умолчанию вы получите 
общее число проверок на пересечение более 500К, что, конечно же, отрицательно скажется на быстродействии.

Чтобы этого избежать, можно использовать флаг ```Sheet::MERGE_NO_CHECK```, тогда проверок не будет, 
и ответственность за то, чтоб адреса не пересекались, полностью лежит на вас:

```php
$sheet->mergeCells('A6:C7', Sheet::MERGE_NO_CHECK);
$sheet->writeTo('G3:I4', 'Hello', null, Sheet::MERGE_NO_CHECK);
```

### Стили текста

Стиль можно задать как для отдельной ячейки, так и для всей строки

```php
$sheet->writeCell('Текст')
    ->applyFontName('Times New Roman')  // название шрифта
    ->applyFontSize(18)                 // размер шрифта 
    ->applyColor('fff000')              // цвет текста
    ->applyBgColor('555555')            // цвет фона
    ;

$sheet->writeRow(['Демо', 'Текст', 'Пример'])
    ->applyFontStyleBold()              // Жирный
    ->applyFontStyleItalic()            // Курсив
    ->applyFontStyleStrikethrough()     // Зачеркнутый текст
    ->applyFontStyleUnderline()         // Подчеркнутый текст
    ;
```
Альтернативный вариант - использовать класс Style для создания массива стилей и передавать полученный массив 
в методы для записи значений.

```php
$style = Excel::newStyle()
    ->setFontName('Times New Roman')    // название шрифта
    ->setFontSize(18)                   // размер шрифта 
    ->setColor('fff000')                // цвет текста
    ->setBgColor('555555')              // цвет фона
    ->setFontStyleBold()
    ->setFontStyleItalic()
;
$sheet->writeCell('Текст', $style);
```

### Выравнивание в ячейке

Выравнивание значения внутри ячейки:
```php
// Текст влево по нижнему краю
$sheet->writeCell('Текст влево')->applyTextAlign(Style::TEXT_ALIGN_LEFT);
// Текст по центру и по горизонтали, и по вертикали
$sheet->writeCell('Текст по центру')->applyTextAlign(Style::TEXT_ALIGN_CENTER, Style::TEXT_ALIGN_CENTER);
// Задаем выравнивание по горизонтали и по вертикали
$sheet->writeCell('Текст вправо')->applyTextAlign(Style::TEXT_ALIGN_RIGHT, Style::TEXT_ALIGN_TOP);
```
### Границы

```php
// Задаем тип и цвет границ (все границы одинаковые
$sheet->writeCell('Текст')->applyBorder(\avadim\FastExcelWriter\Style::BORDER_THIN, '#ff0000');
// Можно для разных границ задать разные опции 
$sheet->writeCell('Текст')
    ->applyBorderTop(\avadim\FastExcelWriter\Style::BORDER_THIN, '#ff0000')
    ->applyBorderRight(\avadim\FastExcelWriter\Style::BORDER_DASHED, '#ff6600')
    ->applyBorderBottom(\avadim\FastExcelWriter\Style::BORDER_DOUBLE, '#000033')
    ->applyBorderLeft(\avadim\FastExcelWriter\Style::BORDER_DASH_DOT, '#336699')
    ;
```
Стили линий:

* BORDER_STYLE_NONE = null;
* BORDER_STYLE_THIN = 'thin';
* BORDER_STYLE_MEDIUM = 'medium';
* BORDER_STYLE_THICK = 'thick';
* BORDER_STYLE_DASH_DOT = 'dashDot';
* BORDER_STYLE_DASH_DOT_DOT = 'dashDotDot';
* BORDER_STYLE_DASHED = 'dashed';
* BORDER_STYLE_DOTTED = 'dotted';
* BORDER_STYLE_DOUBLE = 'double';
* BORDER_STYLE_HAIR = 'hair';
* BORDER_STYLE_MEDIUM_DASH_DOT = 'mediumDashDot';
* BORDER_STYLE_MEDIUM_DASH_DOT_DOT = 'mediumDashDotDot';
* BORDER_STYLE_MEDIUM_DASHED = 'mediumDashed';
* BORDER_STYLE_SLANT_DASH_DOT = 'slantDashDot';

### Изображения

```php
$sheet->addImage($cell, $imageFile, $imageStyle);

// Вставка изображения в ячейку 'A1' из локального файла
$sheet->addImage('A1', 'path/to/file');

// Вставка изображения в ячейку 'A1', скачав его по URL
$sheet->addImage('A1', 'https://site.com/image.jpg');

// Вставка изображения в ячейку 'A1' из строки base64
$sheet->addImage('A1', 'data:image/jpeg;base64,/9j/4AAQ...');

// Вставляем изображение и задаем его ширину в 150 пикселей (высота изменится пропорционально)
$sheet->addImage('B2', 'path/to/file', ['width' => 150]);

// Вставляем изображение и задаем его высоту в 150 пикселей (ширина изменится пропорционально)
$sheet->addImage('C3', 'path/to/file', ['height' => 150]);

// Задаем ширину и высоту вставляемого изображений
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150]);

// Добавляем ссылку к изображению
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150, 'hyperlink' => 'https://www.google.com/']);

```
### Гиперссылки

You can insert URLs as active hyperlinks

```php
// Записываем URL, как обычный текст (без активной ссылки)
$sheet->writeCell('https://google.com');

// Записываем URL, как активную ссылку
$sheet->writeCell('https://google.com', ['hyperlink' => true]);

// Записываем текст как активную ссылку
$sheet->writeCell('Google', ['hyperlink' => 'https://google.com']);
```

Иногда возникает необходимость создать внутреннюю ссылку на другой лист в этой же книге
```php
$sheet->writeCell('Internal link', ['hyperlink' => "Sheet1!C7"]);
```
Если название листа, на который надо сослаться, содержит пробелы, то это название надо взять в одинарные кавычки
```php
$sheet->writeCell('Internal link', ['hyperlink' => "'Sheet 1'!C7"]);
```
Если нужно сослаться на лист в другом XLSX-файле, то имя файла нужно взять в квадратные скобки
```php
$sheet->writeCell('Workbook link', ['hyperlink' => "'[other_file.xlsx]Sheet1'!C7"]);

```
## Формулы

В формулах поддерживаются, как английские, так и русские названия функция (для этого должна быть корректно установлена 
локаль, либо ее надо задать явно ```$sheet->setLocale('ru')``). Оба варианта ниже равнозначны:

```php
$sheet->setCellValue("A3", "=SUM(A1:A2)");
$sheet->setCellValue("A3", "=СУММ(A1:A2)");
```
## Сохранение

```php
// Можно сначала задать имя файла...
$excel->setFileName('/path/to/out/file.xlsx');

// ... а потом сохранить его 
$excel->save();

// А можно явно задать имя файла при сохранении
$excel->save($filename);

// Если требуется отдать файл пользователю на скачивание, то нужно использовать метод ```download()```
$excel->download('name.xlsx');

```
