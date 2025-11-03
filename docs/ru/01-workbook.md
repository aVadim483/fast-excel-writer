## FastExcelWriter — Workbook

### Параметры Книги Excel

```php
// Создание Excel-книги с одним листом по умолчанию 
$excel = Excel::create();

// Создание Excel-книги с листом, который называется 'Abc' 
$excel = Excel::create('Abc');

// Создание Excel-книги с листами 'Foo' и 'Bar'
$excel = Excel::create(['Foo', 'Bar']);

$font = [
    Style::FONT_NAME => 'Arial', 
    Style::FONT_SIZE => 14
];

// Создание Excel-книги и задание шрифта, который будет использоваться по умолчанию
$excel = Excel::create(['Foo', 'Bar'], [Style::FONT => $font]);

// Альтернативный вариант
$excel = Excel::create()->setDefaultStyle([Style::FONT => $font]);
// Или
$excel = Excel::create()->setDefaultFont($font);

// Автоматическая конвертация строк, содержащих числа, в числовые значения
$excel = Excel::create([], ['auto_convert_number' => true]);
$excel->sheet()->writeCell('1234'); // в ячейку будет записано число 1234

// Сохранение строковых значений в общий справочник strings.xml 
$excel = Excel::create([], ['shared_string' => true]);
// Другой способ задать то же
$excel = Excel::create();
$excel->setSharedString();

// Направление письма справа налево
$excel->setRightToLeft(true);

// Задать имя файла по умолчанию
$excel->setFileName('/path/to/out/file.xlsx');

// Сохранение в файл по умолчанию (см. выше) 
$excel->save();

// Сохранение в заданный файл
$excel->save($filename);

// Скачать созданный файл
$excel->download('name.xlsx');
```

### Локаль

Установка локали влияет на форматы, используемые по умолчанию для отображения дат, времени и валюты. 
А так же позволяет использовать названия функций на национальных языках. 

В большинстве случаев локаль устанавливается корректно автоматически, 
но иногда возникает необходимость устанавливать ее вручную

```php
// Задаем французскую локаль
$excel = Excel::create([], ['locale' => 'fr']);
// альтернативный вариант
$excel = Excel::create();
$excel->setLocale('fr');

// Функции на английском языке можно задать в любой локали
$sheet->writeCell('=SUM(A1:C4)');
// Но теперь можно задать ту же функцию и на французском
$sheet->writeCell('=SOMME(A1:C4)');

// Пример использования русской локали
$excel->setLocale('ru');
$sheet->writeCell('=СУММ(A1:C4)');

```

### Метаданные книги

```php
$excel->setMetadata($key, $value);

// Shortcut methods
$excel->setTitle($title);
$excel->setSubject($subject);
$excel->setAuthor($author);
$excel->setCompany($company);
$excel->setDescription($description);
$excel->setKeywords($keywords);

```

### Директория для временных файлов

Библиотека использует временные файлы для генерации XLSX-файла. Если не указано иное, 
они создаются во временной системной директории или в текущей директории выполнения. 
Однако вы можете указать свою директорию для хранения временных файлов.

```php
use \avadim\FastExcelWriter\Excel;

Excel::setTempDir('/path/to/temp/dir'); // используйте этот вызов до Excel::create()
$excel = Excel::create();

// Или альтернативный вариант
$excel = Excel::create('SheetName', ['temp_dir' => '/path/to/temp/dir']);

```
### Справочник строковых значений (Shared Strings)

По умолчанию строки записываются непосредственно в ячейки листа. Это немного увеличивает размер файла,
но ускоряет запись данных и экономит память. Когда вы создаете файл с помощью MS Excel, то строковые значения
сохраняются в общий справочник (strings.xml), а в ячейку записывается ссылка на этот справочник. 
Таким образом уменьшается размер файла, особенно если строки повторяются. 

Если вы хотите воспроизвести это поведение, чтобы строки записывались в общий строковый справочник, 
необходимо использовать параметр 'shared_string'.

```php
$excel = Excel::create([], ['shared_string' => true]);
```

### Вспомогательные методы

Статические вспомогательные методы, которые вы можете использовать в своих приложениях

```php
// Преобразовать букву столбца в порядковый номер (начиная с 1)
$number = Excel::colNumber('C'); // => 3
$number = Excel::colNumber('BZ'); // => 78

// Преобразовать букву столбца в индекс (начиная с 0)
$number = Excel::colIndex('C'); // => 2
$number = Excel::colIndex('BZ'); // => 77

// Обратное преобразование - из порядкового номера в букву
$letter = Excel::colLetter(3); // => 'C'
$letter = Excel::colLetter(78); // => 'BZ'

// Адрес ячейки из номера строки и столбца
$address = Excel::cellAddress(8, 12); // => 'L8'
$address = Excel::cellAddress(8, 12, true); // => '$L$8'

```
