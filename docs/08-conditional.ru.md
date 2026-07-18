## FastExcelWriter – Условное форматирование (с v.6.4)

Условное форматирование позволяет задавать параметры форматирования в зависимости от значения ячейки.
Его можно применять как к отдельным ячейкам, так и к диапазонам ячеек.

### Простое использование

```php
use avadim\FastExcelWriter\Conditional\Conditional;use avadim\FastExcelWriter\Style\Style;

// Ячейкам со значением меньше нуля зададим тёмно-красный цвет текста и светло-красный фон
$cond1 = Conditional::make('<', 0)->setStyle([Style::FONT_COLOR => '#900', Style::FILL_COLOR => '#f99']);
// Ячейкам со значением больше 10 зададим тёмно-зелёный цвет текста и светло-зелёный фон
$cond2 = Conditional::make('>', 10)->setStyle([Style::FONT_COLOR => '#090', Style::FILL_COLOR => '#9f9']);

// Условное форматирование для одной ячейки
$sheet->writeCell(11)
    ->applyConditionalFormatting([$cond1, $cond2]);

// Условное форматирование для колонки C
$sheet->addConditionalFormatting('C1:C' . Excel::MAX_ROW, [$cond1, $cond2]);
```

**ВАЖНО:** если нужно задать условное форматирование для большого количества ячеек, не рекомендуется
использовать функцию `applyConditionalFormatting()` — это может плохо сказаться на производительности.
Используйте `$sheet->addConditionalFormatting($range, $conditional)`, чтобы задать одинаковые условия
форматирования для диапазона ячеек

### Общая проверка значения ячейки

Базовый тип, при котором содержимое ячейки сравнивается с заданным значением.

```php
$value = 10;
$style = [Style::FONT_COLOR => '#900', Style::FILL_COLOR => '#f99'];

// Сравнение с числовыми значениями
$cond = Conditional::make('=', $value, $style);
$cond = Conditional::equals($value, $style); // тот же результат

$cond = Conditional::make('!=', $value, $style);
$cond = Conditional::make('<>', $value, $style);
$cond = Conditional::notEquals($value, $style);

$cond = Conditional::make('>', $value, $style);
$cond = Conditional::greaterThan($value, $style);

$cond = Conditional::make('>=', $value, $style);
$cond = Conditional::greaterThanOrEqual($value, $style);

$cond = Conditional::make('<', $value, $style);
$cond = Conditional::lessThan($value, $style);

$cond = Conditional::make('<=', $value, $style);
$cond = Conditional::lessThanOrEqual($value, $style);

// Оператор "between" требует массива из двух значений.
$cond = Conditional::make('between', [10, 50], $style);
$cond = Conditional::between([10, 50], $style);

$cond = Conditional::make('!between', [10, 50], $style);
$cond = Conditional::notBetween([10, 50], $style);

// Текстовые значения
$cond = Conditional::make('=', 'Hello', $style);
$cond = Conditional::contains('Hello', $style);
$cond = Conditional::notContains('Hello', $style);
$cond = Conditional::beginsWith('Hello', $style);
$cond = Conditional::endsWith('Hello', $style);

// Если вы используете формулы вместо скалярных значений, они должны начинаться с символа "="
$cond = Conditional::make('=', '=B10+SUM(C3:D8)', $style);
```
### Выражения

Если вы используете формулы вместо скалярных значений, они должны начинаться с символа "="

```php
$cond = Conditional::expression('=SUM(C3:D8)>100', $style);
// Адрес RC ссылается на текущую ячейку
$cond = Conditional::expression('=RC>SUM(C3:D8)', $style);
```

### Градиентная заливка по значениям (colorScale)

```php
// Трёхцветная шкала
// -----------------
// Условное форматирование ячеек градиентом от красного через жёлтый к зелёному
// Здесь: ячейки с минимальным значением будут окрашены в красный,
// ячейки с медианным значением (50%) — в жёлтый,
// а ячейки с максимальным значением — в зелёный
$cond = Conditional::colorScale('ff0000', 'fff000', '00ff00');

// Двухцветная шкала
// ---------------
// Условное форматирование ячеек градиентом
// Здесь используются только два цвета — от красного к зелёному
$cond = Conditional::colorScale('ff0000', '00ff00');

// Одноцветная шкала
// ------------------
// Здесь: ячейки с минимальным значением будут окрашены цветом по умолчанию (белым),
// а ячейки с максимальным значением — зелёным
$cond = Conditional::colorScaleMax('00ff00');

// Здесь: ячейки с минимальным значением будут окрашены зелёным,
// а ячейки с максимальным значением — цветом по умолчанию (белым)
$cond = Conditional::colorScaleMin('00ff00');
```
Обычно минимальное и максимальное значения диапазона определяются автоматически,
но их можно задать явно

```php
$cond = Conditional::colorScaleNum([$min, $max], $color1, $color2);
```

### Гистограмма внутри ячейки (dataBar)

Условное форматирование типа dataBar отображает внутри ячеек полосы,
которые визуально представляют значения в определённом диапазоне.

```php
$cond = Conditional::dataBar($color)
    ->gradient(false) // сплошная полоса
    ->showValue(false) // показывать полосу без значений
    ->directionRtl(true) // рисовать полосу справа налево
;
```
