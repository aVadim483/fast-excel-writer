## FastExcelWriter – Книга (Workbook)

### Настройки книги

```php
// Создаёт книгу с одним листом по умолчанию 
$excel = Excel::create();

// Создаёт книгу с одним листом с именем 'Abc' 
$excel = Excel::create('Abc');

// Создаёт книгу с несколькими именованными листами 'Foo' и 'Bar'
$excel = Excel::create(['Foo', 'Bar']);

$font = [
    Style::FONT_NAME => 'Arial', 
    Style::FONT_SIZE => 14
];

// Создаёт книгу со шрифтом по умолчанию
$excel = Excel::create(['Foo', 'Bar'], [Style::FONT => $font]);

// Автоматически преобразовывать строки, содержащие числа, в числа.
// Строки с ведущим нулём ('007') и числа длиннее 15 значащих цифр
// (идентификаторы, штрихкоды) остаются текстом — Excel молча округлил бы их
$excel = Excel::create([], ['auto_convert_number' => true]);

// Сохранять строки в отдельный xml общих строк (shared strings)
$excel = Excel::create([], ['shared_string' => true]);
// или другим способом
$excel = Excel::create();
$excel->setSharedString();

// Задаёт локаль
// В большинстве случаев локаль определяется автоматически,
// но иногда её нужно задать вручную
$excel = Excel::create([], ['locale' => 'fr']);
// или другим способом
$excel = Excel::create();
$excel->setLocale('fr');

// Задаёт форматы книги по умолчанию — они перекрывают форматы локали
// и используются стилями '@date', '@time' и '@datetime'
$excel = Excel::create([], [
    'default_date_format' => 'DD.MM.YYYY',
    'default_time_format' => 'HH:MM',
    'default_datetime_format' => 'DD.MM.YYYY HH:MM',
]);
// или другим способом
$excel = Excel::create();
$excel->setDefaultDateFormat('DD.MM.YYYY');
$excel->setDefaultTimeFormat('HH:MM');
$excel->setDefaultDateTimeFormat('DD.MM.YYYY HH:MM');

// Задаёт шрифт по умолчанию
$excel->setDefaultFont($font);

// Задаёт стили по умолчанию
$excel->setDefaultStyle([Style::FONT => $font]);

// Включает режим «справа налево» (RTL)
$excel->setRightToLeft(true);

// Задаёт имя файла по умолчанию для сохранения
$excel->setFileName('/path/to/out/file.xlsx');

// Сохраняет книгу в файл по умолчанию 
$excel->save();

// Сохраняет книгу в указанный файл 
$excel->save($filename);

// Отдаёт сгенерированный файл клиенту (отправляет в браузер)
$excel->download('name.xlsx');

```

### Класс Options

Вместо массива опций в ```Excel::create()``` можно передать экземпляр класса ```Options``` —
у него есть fluent-интерфейс

```php
use \avadim\FastExcelWriter\Excel;
use \avadim\FastExcelWriter\Options;

$options = Options::create()
    ->tempDir('/path/to/temp/dir') // каталог для временных файлов
    ->tempPrefix('xlsx_') // произвольный префикс временных файлов
    ->autoConvertNumber() // автоматически преобразовывать строки с числами в числа
    ->sharedString() // сохранять строки в xml общих строк
    ->locale('fr') // задать локаль
    ->defaultFont([Style::FONT_NAME => 'Arial', Style::FONT_SIZE => 14]) // задать шрифт по умолчанию
    ->defaultDateFormat('DD.MM.YYYY') // формат стиля '@date'
    ->defaultTimeFormat('HH:MM') // формат стиля '@time'
    ->defaultDateTimeFormat('DD.MM.YYYY HH:MM') // формат стиля '@datetime'
;

$excel = Excel::create(['Sheet1'], $options);
```

См. также: [Класс Options](91-api-class-options.md)

### Метаданные книги

```php
$excel->setMetadata($key, $value);

// Сокращённые методы
$excel->setTitle($title);
$excel->setSubject($subject);
$excel->setAuthor($author);
$excel->setCompany($company);
$excel->setDescription($description);
$excel->setKeywords($keywords);

```

### Каталог для временных файлов

При генерации XLSX-файла библиотека использует временные файлы. Если каталог не указан, они создаются
в системном временном каталоге или в текущем каталоге выполнения. Но вы можете задать каталог для временных файлов явно.

```php
use \avadim\FastExcelWriter\Excel;

Excel::setTempDir('/path/to/temp/dir'); // вызывайте до Excel::create()
$excel = Excel::create();

// Или альтернативный вариант
$excel = Excel::create('SheetName', ['temp_dir' => '/path/to/temp/dir']);

```
### Общие строки (Shared Strings)

По умолчанию строки записываются прямо в листы (*inline strings*). С опцией ```'shared_string'```
каждая строка попадает в отдельную таблицу (```sharedStrings.xml```), а ячейки ссылаются на неё по индексу.

```php
$excel = Excel::create([], ['shared_string' => true]);
```

Выбор здесь — это размен памяти на размер несжатого XML:

* inline-строки пишутся потоково, поэтому память не зависит от объёма данных;
* общие строки требуют держать в памяти таблицу всех уникальных строк до самого сохранения
  (примерно 0,5 КБ на уникальную строку) — это единственная часть записи, которая не потоковая.

Замер на 500 000 строковых ячеек (100 000 строк × 5 столбцов), PHP 8.4:

| данные | режим | xlsx | несжатый XML | пик памяти |
|---|---|---|---|---|
| 5 000 уникальных значений, много повторов | inline | 2,39 МБ | 49,5 МБ | 4 МБ |
| 5 000 уникальных значений, много повторов | shared | 2,09 МБ | 18,9 МБ | 6 МБ |
| все 500 000 значений уникальны | inline | 12,53 МБ | 56,8 МБ | 4 МБ |
| все 500 000 значений уникальны | shared | 12,82 МБ | 59,1 МБ | 284 МБ |

Обратите внимание, как мало меняется размер самого xlsx: ZIP-сжатие и так хорошо сжимает
повторяющиеся строки внутри листа. Общие строки уменьшают *несжатый* XML (Excel открывает такой
файл быстрее), но на уникальных значениях они делают XML даже больше (индекс в ячейке плюс сама строка
в таблице) и обходятся дорого по памяти.

См. также: [Потоковый режим и память](03-writing.md#streaming-mode-and-memory)

**Практическое правило:** для больших выгрузок оставляйте режим по умолчанию. Включайте общие строки,
когда значения берутся из ограниченного словаря (статусы, категории, наименования) и часто повторяются,
либо когда их требует потребитель файла. Библиотека не может решить это за вас: в потоковом режиме
доля повторов неизвестна, пока не записаны все данные.

### Вспомогательные методы

Это статические методы-хелперы, которые можно использовать в ваших приложениях

```php
// Преобразует букву колонки в номер (нумерация с единицы)
$number = Excel::colNumber('C'); // => 3
$number = Excel::colNumber('BZ'); // => 78

// Преобразует букву в индекс (нумерация с нуля)
$number = Excel::colIndex('C'); // => 2
$number = Excel::colIndex('BZ'); // => 77

// Обратное преобразование — из номера в букву (нумерация с единицы)
$letter = Excel::colLetter(3); // => 'C'
$letter = Excel::colLetter(78); // => 'BZ'

// Составляет адрес ячейки из номеров строки и колонки (нумерация с единицы)
$address = Excel::cellAddress(8, 12); // => 'L8'
$address = Excel::cellAddress(8, 12, true); // => '$L$8'
$address = Excel::cellAddress(8, 12, true, false); // => '$L8'
$address = Excel::cellAddress(8, 12, false, true); // => 'L$8'

```
