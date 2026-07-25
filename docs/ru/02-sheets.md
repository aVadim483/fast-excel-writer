## FastExcelWriter – Листы (Sheets)

### Создание, выбор и удаление листа

```php
// Создаёт книгу с тремя именованными листами 
$excel = Excel::create(['Jan', 'Feb', 'Mar']);

// Возвращает первый лист;
$sheet = $excel->getSheet();

// Возвращает лист 'Jan';
$sheet = $excel->getSheet('Jan');

// Возвращает третий лист (с именем 'Mar');
$sheet = $excel->getSheet(3);

// Создаёт новый лист с именем 'Total'
$sheet = $excel->makeSheet('Total');

$sheet->setName($sheetName);

// Удаляет указанный лист
$excel->removeSheet('Total');
```

### Настройки листа

```php
$sheet->setDefaultFont($font);
$sheet->setDefaultFontName($fontName);
$sheet->setDefaultFontSize($fontSize);
$sheet->setDefaultFontStyle($fontStyle);
$sheet->setDefaultFontStyleBold();
$sheet->setDefaultFontStyleItalic();
$sheet->setDefaultFontStyleUnderline(true);
$sheet->setDefaultFontStyleStrikethrough();
$sheet->setDefaultFontColor($font);

$sheet->setTabColor('#ff0099');

$sheet->setStateHidden();
$sheet->setStateVeryHidden();
$sheet->setStateVisible();
// ИЛИ
$sheet->setSheetState('hidden'); // или 'visible', 'veryHidden'
```


### Закрепление областей и автофильтр

```php
$sheet->setFreeze('B2');
$sheet->setAutofilter(1);
```

### Настройки строк

Параметры строки (стили и высоту) можно задать вторым аргументом функции ```writeRow()```.
Обратите внимание, что в этом случае стили будут применены только к тем ячейкам строки, в которые записаны данные

```php
// Записываем строку данных и задаём высоту
$rowStyle = [
    'fill-color' => '#fffeee',
    'border' => 'thin',
    'height' => 28,
];
$sheet->writeRow(['aaa', 'bbb', 'ccc'], $rowStyle);
```
Другой способ с тем же результатом

```php
$sheet->writeRow(['aaa', 'bbb', 'ccc', null, 'eee'])
    ->applyFillColor('#fffeee')
    ->applyBorder('thin')
    ->applyRowHeight(28);

```
Можно задать высоту строки или её видимость

```php
// Задаём высоту строки 2 равной 33
$sheet->setRowHeight(2, 33);

// Задаём высоту строк 3, 5 и 7 равной 33
$sheet->setRowHeight([3, 5, 7], 33);

// Задаём высоту нескольких строк
$sheet->setRowHeights([1 => 20, 2 => 33, 3 => 40]);

// Скрываем строку 8
$sheet->setRowVisible(8, false);

// Другой способ
$sheet->setRowHidden(8);

// Скрываем строки 9, 10, 11
$sheet->setRowVisible([9, 10, 11], false);

// Показываем строку 10
$sheet->setRowVisible(10, true);
```
ВАЖНО: функции setRowXX() можно применять только к строкам с номером не меньше текущего.
См. [Построчная и прямая запись](03-writing.md#writing-row-by-row-vs-direct).
Поэтому следующий код выбросит ошибку «Row number must be greater than written rows»

```php
$sheet = $excel->sheet();
// Записываем строку 1
$sheet->writeRow(['aaa1', 'bbb1', 'ccc1']);
// Записываем строку 2
$sheet->writeRow(['aaa2', 'bbb2', 'ccc2']);
// Пытаемся задать высоту уже записанной строки 1
$sheet->setRowHeight(1, 33);

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

### Настройки колонок

Ширину колонок можно задать несколькими способами

```php
// Задаём ширину колонки D равной 24
$sheet->setColWidth('D', 24);
$sheet->setColDataStyle('D', ['width' => 24]);

// Задаём ширину конкретных колонок
$sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);
// Задаём ширину колонок начиная с 'A'
$sheet->setColWidths([10, 20, 30, 40], 24);

$colStyle = [
    'B' => ['width' => 10], 
    'C' => ['width' => 'auto'], 
    'E' => ['width' => 30], 
    'F' => ['width' => 40],
];
$sheet->setColDataStyleArray($colStyle);

```
Можно задать минимальную ширину колонок. Обратите внимание, что минимальное значение имеет более высокий приоритет
```php
// Задаём минимальную ширину 20 
$sheet->setColMinWidth('D', 20);

// Значение 10 не будет установлено, потому что оно меньше минимального
$sheet->setColWidth('D', 10);

// А ширина 30 будет установлена
$sheet->setColWidth('D', 30);

// Ширина колонки будет подобрана по содержимому, но не меньше 20
$sheet->setColWidthAuto('D');

// Скрываем колонку B
$sheet->setColVisible('B', false);

// Другой способ
$sheet->setColHidden('B');

// Скрываем колонки B, E, H
$sheet->setColVisible(['B', 'E', 'H'], false);

// Показываем колонку E
$sheet->setColVisible('E', true);
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

### Автоматическая ширина колонок

```php
// Задаём автоматическую ширину
$sheet->setColWidth('D', 'auto');
$sheet->setColWidthAuto('D');
$sheet->setColDataStyle('D', ['width' => 'auto']);

// Задаём ширину конкретных колонок
$sheet->setColWidths(['B' => 10, 'C' => 'auto', 'E' => 30, 'F' => 40]);

$colStyle = [
    'B' => ['width' => 10], 
    'C' => ['width' => 'auto'], 
    'E' => ['width' => 30], 
    'F' => ['width' => 40],
];
$sheet->setColDataStyleArray($colStyle);
```

**ВАЖНО!** Автоматический расчёт ширины колонки — очень сложная задача. Используя эту опцию, учитывайте следующее:

1. Расчёт выполняется для каждой ячейки колонки по мере заполнения листа.
Поэтому автоподбор ширины колонки должен быть включён до того, как вы начнёте записывать данные в ячейки.
Если включить его в конце создания документа, он не даст никакого эффекта.

2. Расчёт выполняется весьма приблизительно, исходя из настроек шрифта и количества символов в значении ячейки.
Но ширина разных символов различается, поэтому ширина колонки может оказаться как больше ширины текста в ячейке,
так и меньше.

3. Библиотека не умеет вычислять формулы, поэтому если ячейка содержит формулу, она игнорируется при расчёте
автоматической ширины колонки.

### Группировка строк и колонок

Задаём уровень группировки для указанных строк

```php
$sheet = $excel->sheet();

// первый уровень
$sheet->writeRow($rowData1)->applyRowOutlineLevel(1);
$sheet->writeRow($rowData2)->applyRowOutlineLevel(1);

// второй уровень
$sheet->writeRow($rowData3)->applyRowOutlineLevel(2);
$sheet->writeRow($rowData4)->applyRowOutlineLevel(2);

// возвращаемся на первый уровень
$sheet->writeRow($rowData5)->applyRowOutlineLevel(1);

// записываем строки без группировки
$sheet->writeRow($rowData6);
$sheet->writeRow($rowData7);
```

Можно настроить группировку для будущих строк.

```php
// уровень 1 для строки 4
$sheet->setRowOutlineLevel(4, 1);

// уровень 1 для строк 5, 6, 7
$sheet->setRowOutlineLevel([5, 6, 7], 1);

// уровень 1 для строк с 9 по 15
$sheet->setRowOutlineLevel('9:15', 1);
// уровень 2 для строк с 11 по 13
$sheet->setRowOutlineLevel('11:13', 2);
```

Можно настроить группировку для последовательности строк.

```php
$sheet = $excel->sheet();

// Записываем строки без группировки
$sheet->writeRow([...]);
$sheet->writeRow([...]);

// Повышаем уровень группировки (устанавливаем уровень 1)
$sheet->beginOutlineLevel();
$sheet->writeRow([...]);
$sheet->writeRow([...]);

// Снова повышаем уровень группировки (уровень 2) со сворачиванием
$sheet->beginOutlineLevel(true);
$sheet->writeRow([...]);
$sheet
    ->writeCell('...')
    ->writeCell('...')
    ->writeCell('...')
    ->nextRow();
$sheet->writeRow([...]);

// Понижаем уровень группировки (обратно к 1)
$sheet->endOutlineLevel();
$sheet->writeRow([...]);

// Возвращаемся к нулевому уровню
$sheet->endOutlineLevel();
```

Задаём уровень группировки для указанных колонок

```php
$sheet->setColOutlineLevel('B', 1);
$sheet->setColOutlineLevel('C', 1);
$sheet->setColOutlineLevel('D', 1);

$sheet->setColOutlineLevel(['F', 'g', 'h', 'i', 'J'], 1);
$sheet->setColOutlineLevel('g:i', 2);

```

### Именованные диапазоны

FastExcelWriter поддерживает _именованные диапазоны_ и не поддерживает _именованные формулы_.
_Именованный диапазон_ задаёт имя-ссылку на ячейку или диапазон ячеек.
Все _именованные диапазоны_ добавляются в книгу, поэтому имена должны быть уникальными,
но определять _именованные диапазоны_ можно как на листе, так и в книге.

Имена диапазонов должны начинаться с буквы или подчёркивания, не содержать пробелов и быть не длиннее 255 символов.

```php
$excel = Excel::create();
$excel->setFileName($outFileName);
$sheet = $excel->sheet();

// Именуем одну ячейку
$sheet->addNamedRange('B2', 'cell_name');

// Именованный диапазон на листе
$sheet->addNamedRange('c2:e3', 'range_name');

// Добавляем именованный диапазон в книгу (требуется имя листа)
$excel->addNamedRange('Sheet1!A1:F5', 'A1_F5');

// Имя можно задать с помощью applyNamedRange()
$sheet->writeCell(1000)->applyNamedRange('Value');
$sheet->writeCell(0.12)->applyNamedRange('Rate');
// Добавляем формулу, использующую имена
$sheet->writeCell('=Value*Rate');

```

### Активный лист и активные ячейки

Можно выбрать активный (открываемый по умолчанию) лист книги

```php
// Задаёт активный лист по имени (без учёта регистра)
$excel->setActiveSheet($name);
```

Чтобы выбрать активную ячейку на указанном листе, используйте следующий код:

```php
// Выбираем одну активную ячейку
$sheet->setActiveCell('B2');

// Выбираем диапазон ячеек
$sheet->setActiveCell('B2:C3');
```

### Параметры страницы { #page-setup }

```php
// Задаём размер бумаги
$sheet->pagePaperSizeA4();
$sheet->pagePaperSizeA3();
$sheet->pagePaperSizeLetter();
$sheet->pagePaperSizeLegal();

// Другие размеры определены в константах Excel::PAPERSIZE_*
$sheet->pagePaperSize(Excel::PAPERSIZE_A4);
$sheet->pagePaperSize(Excel::PAPERSIZE_B5);
$sheet->pagePaperSize(Excel::PAPERSIZE_JAPANESE_POSTCARD_ROTATED);

$sheet->pagePaperHeight('297mm');
$sheet->pagePaperWidth('21cm');

// Масштабирование страницы
$sheet->pageScale(100); // 100%
$sheet->pageFitToWidth(1); // вписать по ширине в 1 страницу
$sheet->pageFitToHeight(0);// автоподбор по высоте

// Задаём ориентацию страницы
$sheet->pagePortrait();
$sheet->pageLandscape();

// Задаём поля
$sheet->pageMargins([
        'left' => '0.5',
        'right' => '0.5',
        'top' => '1.0',
        'bottom' => '1.0',
        'header' => '0.5',
        'footer' => '0.5',
    ]);
// то же самое действие    
$sheet
    ->pageMarginLeft(0.5)
    ->pageMarginRight(0.5)
    ->pageMarginTop(1.0)
    ->pageMarginBottom(1.0)
    ->pageMarginHeader(0.5)
    ->pageMarginFooter(0.5);

```
По умолчанию значения задаются в дюймах, 1 дюйм равен 2,54 см. То есть числовые значения интерпретируются как дюймы.

Но эти значения можно указывать и в сантиметрах или миллиметрах.

```php
$sheet->pageMarginLeft(0.5); // левое поле 0.5 дюйма
$sheet->pageMarginLeft('0.5cm'); // левое поле 0.5 сантиметра
$sheet->pageMarginLeft('0.5mm'); // левое поле 0.5 миллиметра
```

### Настройки печати

Задаём область печати

```php
$sheet->setPrintArea('A2:F3');
$sheet->setPrintArea('A8:F10');

// Или несколько областей печати одним вызовом
$sheet->setPrintArea('A2:F3,A8:F10');
```

Чтобы повторять определённые строки/колонки сверху/слева на каждой печатной странице, используйте следующий код:

```php
$sheet->setPrintTopRows('1')->setPrintLeftColumns('A');
```

Пример повторения строк с 1 по 5 и колонок A–C на каждой печатной странице:

```php
$sheet->setPrintTopRows('1:5');
$sheet->setPrintLeftColumns('A:C');
```

Чтобы показать/скрыть линии сетки при печати, используйте следующий код:

```php
$sheet->setPrintGridlines(true);
```

Область печати можно центрировать на странице

```php
// центрирование по горизонтали (параметр по умолчанию — true)
$sheet->setPrintHorizontalCentered();

// центрирование по вертикали
$sheet->setPrintVerticalCentered();

// центрирование в обоих направлениях
$sheet->setPrintCentered();
```
### Колонтитулы при печати

```php
// Задаём одинаковые верхний и нижний колонтитулы для всех страниц
$sheet->pageHeaderFooter('Print Header', 'Print Footer');

$sheet->pageHeader('Header for all pages');
$sheet->pageHeaderFirst('Header for the First page');
$sheet->pageHeaderOdd('Header for Odd pages');
$sheet->pageHeaderEven('Header for Even pages');

$sheet->pageFooter('Footer for all pages');
$sheet->pageFooterFirst('Footer for the First page');
$sheet->pageFooterOdd('Footer for Odd pages');
$sheet->pageFooterEven('Footer for Even pages');
```

При определении колонтитулов можно использовать специальные коды форматирования, начинающиеся с `&`.  
Ниже приведён полный практический список.

#### Поля страницы и документа

- **`&P`** — номер текущей страницы
- **`&N`** — общее количество страниц
- **`&D`** — текущая дата
- **`&T`** — текущее время
- **`&A`** — имя листа
- **`&F`** — имя файла
- **`&Z`** — путь к файлу
- **`&G`** — вставка изображения (картинки в колонтитуле)

```php
$sheet->pageFooter('Page &P of &N');
```

#### Шрифт и форматирование текста

- **`&"FontName,Style"`** — задать шрифт и начертание
- **`&12`** — задать размер шрифта (число = размер)
```php
$sheet->pageFooter('&"Arial,Bold Italic"Page &P of &N');
```

Переключатели начертания:

- **`&B`** — полужирный
- **`&I`** — курсив
- **`&U`** — подчёркивание
- **`&E`** — двойное подчёркивание
- **`&S`** — зачёркивание
- **`&X`** — надстрочный текст
- **`&Y`** — подстрочный текст

Цвет:

- **`&Krrggbb`** — цвет текста (hex RGB)

#### Позиционирование внутри колонтитула

Эти коды определяют выравнивание текста внутри колонтитула:

- **`&L`** — левая секция
- **`&C`** — центральная секция
- **`&R`** — правая секция

```php
$sheet->pageHeaderFirst('&LLeft &CCenter &RRight');
```

#### Дополнительные правила

- **`&&`** — выводит символ `&` (иначе он трактуется как управляющий код)

Коды можно свободно комбинировать
