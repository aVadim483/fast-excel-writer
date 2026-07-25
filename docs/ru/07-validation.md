## FastExcelWriter – Проверка данных (с v.6.0)

Проверка данных (data validation) позволяет установить входной фильтр на данные, которые можно ввести в конкретную ячейку.
Библиотека позволяет задать следующие типы фильтров:

* integer (целое число)
* decimal (число)
* date (дата)
* text length (длина текста)
* dropdown (список)
* custom (пользовательский)

Со всеми типами фильтров (кроме "dropdown" и "custom") можно использовать следующие операторы:

* равно ('=')
* не равно ('!=')
* между (between)
* не между (not between)
* больше ('>')
* больше или равно ('>=')
* меньше ('<')
* меньше или равно ('<=')

### Простое использование

```php
use avadim\FastExcelWriter\DataValidation\DataValidation;

$sheet->writeCell('Integer:');
// Значение следующей ячейки должно быть целым числом от 1 до 9
$sheet->nextCell()->applyDataValidation(DataValidation::integer('between', [1, 9]));

// Другой способ
$validation = DataValidation::decimal('>', '=B5');
$sheet->nextRow();
$sheet->writeTo('B5', 12.34);
$sheet->writeCell('Decimal:');
// Значение следующей ячейки должно быть числом (float), большим значения B5
$sheet->nextCell()->applyDataValidation($validation);

```

### Определение фильтров

#### DataValidation::integer($operator, $formulas);

```$operator``` — строка. Доступные операторы: '=', '!=', 'between', '!between', '>', '>=', '<', '<='.
Также можно использовать константы ```DataValidation::OPERATOR_*```

```$formulas``` может быть числом или строкой. Для операторов 'between' и '!between' ```$formulas``` должен быть массивом чисел или строк.

Есть три способа задать формулу:

1. Просто число 
```php
$validation = DataValidation::decimal('>', 123);
$validation = DataValidation::decimal('>', '123');
```
2. Ссылка на другую ячейку
```php
$validation = DataValidation::decimal('>', '=B48');
```
3. Через формулу Excel
```php
$validation = DataValidation::decimal('>', '=SUM(A2:A10)+D18');
```

Если оператор — "between" или "!between", второй аргумент должен быть массивом из двух значений/формул.
```php
$validation = DataValidation::decimal('!between', [-1, '=A5-D6']);
```

#### DataValidation::decimal($operator, $formulas);

Аналогично ```integer```.

#### DataValidation::textLength($operator, $formulas);

Аналогично ```integer```.

#### DataValidation::date($operator, $formulas);

Используются те же операторы, что и в ```integer``` или ```decimal```.
Но если вы хотите использовать скалярные значения в качестве формул, они должны быть метками времени (timestamp).
```php
$validation = DataValidation::date('>', Excel::toTimestamp('2024-01-01'));
``` 

#### DataValidation::dropDown($formulas);

```php
// Задаём выпадающий список
$validation = DataValidation::dropDown(['item1', 'item2', 'item3']);
// Берём элементы из диапазона
$validation = DataValidation::dropDown('=A1:A5');
// Берём элементы из именованного диапазона
$validation = DataValidation::dropDown('=sheet1!list');
``` 

### Проверка типа значения

```php
$validation = DataValidation::isNumber();
$validation = DataValidation::isText();
```

### Пользовательские фильтры
В следующем примере значение в ячейке должно начинаться с префикса "ID-" и быть длиной не менее 10 символов.
```php
$validation = DataValidation::custom('=AND(LEFT(RC,3)="ID-", LEN(RC)>9)');
```
Обратите внимание, что для ссылки на текущую ячейку используется адрес "RC".


### Все настройки проверки данных

```php
$validation = DataValidation::make(DataValidation::TYPE_INTEGER);
$validation
    ->setOperator('between')
    ->setFormula1('=F23')
    ->setFormula2(43)
    ->allowBlank() // разрешить пустое значение
    ->setErrorStyle() // stop, warning или information
    ->setError($errorMessage, $errorTitle)
    ->setPrompt($promptMessage, $promptTitle)
;

$sheet->addDataValidation('E32', $validation);
```

Другие методы

```php
// Разрешить пустое значение 
$validation->allowBlank();

// Запретить пустое значение 
$validation->allowBlank(false);

// Показывать выпадающий список
$validation->showDropDown();

// Не показывать выпадающий список 
$validation->showDropDown(false);

// Показывать подсказку при вводе
$validation->showInputMessage();

// Не показывать подсказку при вводе 
$validation->showInputMessage(false);

// Показывать сообщение об ошибке
$validation->showErrorMessage();

// Не показывать сообщение об ошибке
$validation->showErrorMessage(false);
```

### Более 64K правил проверки

ВАЖНО! Нельзя задавать более 64K правил проверки — это может вызвать ошибку при открытии файла в Excel.

Если нужно задать проверку данных для определённой области, используйте следующий код:
```php
$sheet = $excel->sheet();
foreach ($someDataArray as $rowData) {
    // здесь записываем данные на лист
}
$validation = DataValidation::dropDown(['item1', 'item2', 'item3']);
$sheet->addDataValidation('B10:E32', $validation);
```
То есть сначала заполните лист данными, а затем вызовите этот метод.
