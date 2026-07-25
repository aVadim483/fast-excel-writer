# FastExcelWriter – Руководство по обновлению

На этой странице описаны важные изменения между основными версиями, которые могут потребовать
обновления вашего кода.

## Обновление до версии 6

Главная новость v.6.0 — поддержка [проверки данных](07-validation.md).

### Важные изменения в версии 6.1

* ```Sheet::setRowOptions()```, ```Sheet::setColOptions()```, ```Sheet::setRowStyles()``` и ```Sheet::setColStyles()```
объявлены устаревшими, вместо них следует использовать другие функции: ```setRowStyle()```, ```setRowStyleArray()```,
```setRowDataStyle()```, ```setRowDataStyleArray()```, ```setColStyle()```, ```setColStyleArray()```, ```setColDataStyle()```, ```setColDataStyleArray()```
* Изменилось поведение ```Sheet::setRowStyle()``` и ```Sheet::setColStyle()``` — теперь они задают стили
для всей строки или колонки (даже для пустых ячеек)

### Важные изменения в версии 6.9

* Namespace класса ```RichText``` изменён на ```avadim\FastExcelWriter\RichText```
* Namespace классов ```Style```, ```StyleManager``` и ```Font``` изменён на ```avadim\FastExcelWriter\Style```
* Удалены устаревшие методы: ```Sheet::setColStyles()```, ```Sheet::setColOptions()```, ```Sheet::getExternalLinks()```,
```Sheet::setPageOptions()```, ```Sheet::setRowOptions()```, ```Sheet::setRowStyles()```

## Обновление до версии 5

Главная новость v.5.0 — поддержка [диаграмм](05-charts.md).

### Важные изменения в версии 5.8

До v.5.8

```php
$sheet->writeCell(12345); // В ячейку будет записано число 12345
$sheet->writeCell('12345'); // Здесь тоже будет записано число 12345
```

В версии 5.8 и позже

```php
$sheet->writeCell(12345); // В ячейку будет записано число 12345
$sheet->writeCell('12345'); // Здесь в ячейку будет записана строка '12345'
```

Если вы хотите сохранить прежнее поведение для обратной совместимости,
используйте опцию 'auto_convert_number' при создании книги.

```php
$excel = Excel::create(['Sheet1'], ['auto_convert_number' => true]);
$sheet = $excel->sheet();
$sheet->writeCell('12345'); // Строка '12345' будет автоматически преобразована в число
```

### Переименованные методы в версии 5.1

Некоторые методы диаграмм были переименованы

* ```setDataSeriesTickLabels()``` => ```setCategoryAxisLabels()```
* ```setXAxisLabel()``` => ```setCategoryAxisTitle()```
* ```getXAxisLabel()``` => ```getCategoryAxisTitle()```
* ```setYAxisLabel()``` => ```setValueAxisTitle()```
* ```getYAxisLabel()``` => ```getValueAxisTitle()```

## Обновление до версии 4

* Библиотека стала работать ещё быстрее
* Добавлен fluent-интерфейс для применения стилей
* Новые методы и рефакторинг кода

Полный список изменений см. в [CHANGELOG](https://github.com/aVadim483/fast-excel-writer/blob/master/CHANGELOG.md).
