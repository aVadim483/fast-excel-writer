## FastExcelWriter - Data Validation (since v.6.0)

Data validation allows to specify an input filter on the data that can be inserted in a specific cell.
The library allows you to set the following types of filters:

* integer (whole)
* decimal
* date
* text length
* dropdown (list)
* custom

The following operators can be used with all filter types (except "dropdown" and "custom"):

* equal ('=')
* not equal ('!=')
* between
* not between
* greater than ('>')
* greater than or equal ('>=')
* less than ('<')
* less than or equal ('<=')

### Simple usage

```php
use avadim\FastExcelWriter\DataValidation\DataValidation;

$sheet->writeCell('Integer:');
// Value of the next cell must be integer between 1 and 9
$sheet->nextCell()->applyDataValidation(DataValidation::integer('between', [1, 9]));

// Other way
$validation = DataValidation::decimal('>', '=B5');
$sheet->nextRow();
$sheet->writeTo('B5', 12.34);
$sheet->writeCell('Decimal:');
// Value of the next cell must be decimal (float) and greater than value of B5
$sheet->nextCell()->applyDataValidation($validation);

```

### Define filters

#### DataValidation::integer($operator, $formulas);

```$operator``` is string. Available operators: '=', '!=', 'between', '!between', '>', '>=', '<', '<='.
Also, you can use constants ```DataValidation::OPERATOR_*```

```$formulas``` can be a number or a string. For operators 'between', '!between' $formula must be an array of numbers or strings.

There are three ways to set up a formula:

1. Just number 
```php
$validation = DataValidation::decimal('>', 123);
$validation = DataValidation::decimal('>', '123');
```
2. Link to other cell
```php
$validation = DataValidation::decimal('>', '=B48');
```
3. Via Excel formula
```php
$validation = DataValidation::decimal('>', '=SUM(A2:A10)+D18');
```

If the operator is "between" or "!between", then the second argument must be an array of two elements of values/formulas.
```php
$validation = DataValidation::decimal('!between', [-1, '=A5-D6']);
```

#### DataValidation::decimal($operator, $formulas);

The same as ```integer```.

#### DataValidation::textLength($operator, $formulas);

The same as ```integer```.

#### DataValidation::date($operator, $formulas);

The same operators are used as in ```integer``` or ```decimal```.
But if you want to use scalar values as formulas, then they must be timestamps.
```php
$validation = DataValidation::date('>', Excel::toTimestamp('2024-01-01'));
``` 

#### DataValidation::dropDown($formulas);

```php
// Set dropdown list
$validation = DataValidation::dropDown(['item1', 'item2', 'item3']);
// Get items from range
$validation = DataValidation::dropDown('=A1:A5');
// Get items from named range
$validation = DataValidation::dropDown('=sheet1!list');
``` 

### Check type of value

```php
$validation = DataValidation::isNumber();
$validation = DataValidation::isText();
```

### Custom filters
In the following example, the value in the cell must begin with the prefix "ID-" and be at least 10 characters long.
```php
$validation = DataValidation::custom('=AND(LEFT(RC,3)="ID-", LEN(RC)>9)');
```
Note that the address "RC" is used to reference the current cell.


### All Data Validation settings

```php
$validation = DataValidation::make(DataValidation::TYPE_INTEGER);
$validation
    ->setOperator('between')
    ->setFormula1('=F23')
    ->setFormula2(43)
    ->allowBlank() // allow blank value
    ->setErrorStyle() // stop, warning or information
    ->setError($errorMessage, $errorTitle)
    ->setPrompt($promptMessage, $promptTitle)
;

$sheet->addDataValidation('E32', $validation);
```

Other methods

```php
// Allow blank value 
$validation->allowBlank();

// Disallow blank value 
$validation->allowBlank(false);

// Show dropdown list
$validation->showDropDown();

// Disallow dropdown list 
$validation->showDropDown(false);

// Show input message
$validation->showInputMessage();

// Disallow input message 
$validation->showInputMessage(false);

// Show error message
$validation->showErrorMessage();

// Disallow error message
$validation->showErrorMessage(false);
```
