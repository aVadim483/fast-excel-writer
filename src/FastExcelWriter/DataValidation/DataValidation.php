<?php

namespace avadim\FastExcelWriter\DataValidation;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Exceptions\ExceptionDataValidation;
use avadim\FastExcelWriter\Sheet;

class DataValidation
{
    /* Data validation types */
    const TYPE_CUSTOM = 'custom'; // Data validation which uses a custom formula to check the cell value
    const TYPE_DATE = 'date'; // Data validation which checks for date values satisfying the given condition
    const TYPE_DECIMAL = 'decimal'; // Data validation which checks for decimal values satisfying the given condition
    const TYPE_LIST = 'list'; // Data validation which checks for a value matching one of list of values
    const TYPE_NONE = 'none'; // No data validation. Why do we need this type?
    const TYPE_TEXT_LENGTH = 'textLength'; // Data validation which checks for text values, whose length satisfies the given condition
    const TYPE_TEXTLENGTH = self::TYPE_TEXT_LENGTH;
    const TYPE_TIME = 'time'; // Data validation which checks for time values satisfying the given condition
    const TYPE_WHOLE = 'whole'; // Data validation which checks for whole number values satisfying the given condition
    const TYPE_INTEGER = self::TYPE_WHOLE;


    /* Data validation error styles */
    const STYLE_STOP = 'stop';
    const STYLE_WARNING = 'warning';
    const STYLE_INFORMATION = 'information';


    /* Data validation operators */
    const OPERATOR_BETWEEN = 'between';
    const OPERATOR_EQUAL = 'equal';
    const OPERATOR_GREATER_THAN = 'greaterThan';
    const OPERATOR_GREATERTHAN = self::OPERATOR_GREATER_THAN;
    const OPERATOR_GREATER_THAN_OR_EQUAL = 'greaterThanOrEqual';
    const OPERATOR_GREATERTHANOREQUAL = self::OPERATOR_GREATER_THAN_OR_EQUAL;
    const OPERATOR_LESS_THAN = 'lessThan';
    const OPERATOR_LESSTHAN = self::OPERATOR_LESS_THAN;
    const OPERATOR_LESS_THAN_OR_EQUAL = 'lessThanOrEqual';
    const OPERATOR_LESSTHANOREQUAL = self::OPERATOR_LESS_THAN_OR_EQUAL;
    const OPERATOR_NOT_BETWEEN = 'notBetween';
    const OPERATOR_NOTBETWEEN = self::OPERATOR_NOT_BETWEEN;
    const OPERATOR_NOT_EQUAL = 'notEqual';
    const OPERATOR_NOTEQUAL = self::OPERATOR_NOT_EQUAL;


    protected string $type;

    protected Sheet $sheet;
    protected string $sqref;

    protected ?string $errorStyle = null;

    protected ?string $operator = null;
    protected ?int $allowBlank = 1;
    protected ?string $showDropDown = null;
    protected ?int $showInputMessage = null;
    protected ?int $showErrorMessage = null;
    protected ?string $errorTitle = null;
    protected ?string $error = null;
    protected ?string $promptTitle = null;
    protected ?string $prompt = null;
    protected ?string $formula1 = null;
    protected ?string $formula2 = null;

    protected array $aliases = [
        '=' => self::OPERATOR_EQUAL,
        '!=' => self::OPERATOR_NOT_EQUAL,
        '>' => self::OPERATOR_GREATER_THAN,
        '>=' => self::OPERATOR_GREATER_THAN_OR_EQUAL,
        '<' => self::OPERATOR_LESS_THAN,
        '<=' => self::OPERATOR_LESS_THAN_OR_EQUAL,
        '!between' => self::OPERATOR_NOT_BETWEEN,
    ];

    protected array $availableTypes = [
        self::TYPE_CUSTOM,
        self::TYPE_DATE,
        self::TYPE_DECIMAL,
        self::TYPE_LIST,
        self::TYPE_NONE,
        self::TYPE_TEXT_LENGTH,
        self::TYPE_TIME,
        self::TYPE_WHOLE,
    ];

    protected array $availableOperators = [
        self::OPERATOR_BETWEEN,
        self::OPERATOR_EQUAL,
        self::OPERATOR_GREATER_THAN,
        self::OPERATOR_GREATER_THAN_OR_EQUAL,
        self::OPERATOR_LESS_THAN,
        self::OPERATOR_LESS_THAN_OR_EQUAL,
        self::OPERATOR_NOT_BETWEEN,
        self::OPERATOR_NOT_EQUAL,
    ];

    protected array $availableErrorStyles = [
        self::STYLE_STOP,
        self::STYLE_WARNING,
        self::STYLE_INFORMATION,
    ];

    /**
     * DataValidation constructor
     *
     * @param $type
     */
    public function __construct($type)
    {
        if (!in_array($type, $this->availableTypes)) {
            ExceptionDataValidation::throwNew('Invalid type for data validation "' . $type . '"');
        }

        $this->type = $type;
    }

    /**
     * Make a DataValidation instance
     *
     * @param $type
     *
     * @return DataValidation
     */
    public static function make($type): DataValidation
    {
        return new self($type);
    }

    /**
     * Make data validation as an integer value
     *
     * @param string $operator
     * @param string|int|array $formulas
     *
     * @return DataValidation
     */
    public static function integer(string $operator, $formulas): DataValidation
    {
        $validation = new self(self::TYPE_WHOLE);
        $validation->setOperator($operator, $formulas);

        return $validation;
    }

    /**
     * Alias of integer()
     *
     * @param string $operator
     * @param string|int|array $formulas
     *
     * @return DataValidation
     */
    public static function whole(string $operator, $formulas): DataValidation
    {

        return self::integer($operator, $formulas);
    }

    /**
     * Make data validation as a decimal value
     *
     * @param string $operator
     * @param string|int|array $formulas
     *
     * @return DataValidation
     */
    public static function decimal(string $operator, $formulas): DataValidation
    {
        $validation = new self(self::TYPE_DECIMAL);
        $validation->setOperator($operator, $formulas);

        return $validation;
    }

    /**
     * Make data validation as a date value
     *
     * @param string $operator
     * @param string|int|array $formulas
     *
     * @return DataValidation
     */
    public static function date(string $operator, $formulas): DataValidation
    {
        $validation = new self(self::TYPE_DATE);
        $validation->setOperator($operator, $formulas);

        return $validation;
    }

    /**
     * Make data validation as a dropdown list
     *
     * @param array|string $formulas
     *
     * @return DataValidation
     */
    public static function dropDown($formulas): DataValidation
    {
        $validation = new self(self::TYPE_LIST);
        $validation->setFormula1($formulas);

        return $validation;
    }

    /**
     * Alias of dropDown()
     *
     * @param array|string $formulas
     * *
     * @return DataValidation
     */
    public static function list($formulas): DataValidation
    {

        return self::dropDown($formulas);
    }

    /**
     * Make data validation as a text length
     *
     * @param string $operator
     * @param array|string $formulas
     *
     * @return DataValidation
     */
    public static function textLength(string $operator, $formulas): DataValidation
    {
        $validation = new self(self::TYPE_TEXT_LENGTH);
        $validation->setOperator($operator, $formulas);

        return $validation;
    }

    /**
     * Make data validation as a custom rule
     *
     * @param string $formula
     *
     * @return DataValidation
     */
    public static function custom(string $formula): DataValidation
    {
        $validation = new self(self::TYPE_CUSTOM);
        $validation->setFormula1($formula);

        return $validation;
    }

    /**
     * Make data validation as an expression (alias of self::custom())
     *
     * @param string $formula
     *
     * @return DataValidation
     */
    public static function expression(string $formula): DataValidation
    {

        return self::custom($formula);
    }

    /**
     * Checking if a cell value is a number
     *
     * @return DataValidation
     */
    public static function isNumber(): DataValidation
    {

        return self::custom('=ISNUMBER(RC)');
    }

    /**
     * Checking if a cell value is a text
     *
     * @return DataValidation
     */
    public static function isText(): DataValidation
    {

        return self::custom('=ISTEXT(RC)');
    }

    /**
     * @param $formula
     *
     * @return string
     */
    protected function checkFormula($formula): string
    {
        if (is_array($formula)) {
            $result = '"' . implode(',', $formula) . '"';
        }
        elseif (is_float($formula)) {
            $result = str_replace(',' , '.', (string)$formula);
        }
        elseif (is_bool($formula)) {
            $result = (int)$formula;
        }
        else {
            $result = (string)$formula;
        }

        return $result;
    }

    /**
     * Set formula 1 for data validation
     *
     * @param int|float|string|array|null $formula
     *
     * @return $this
     */
    public function setFormula1($formula): DataValidation
    {
        $this->formula1 = ($formula !== null) ? $this->checkFormula($formula) : null;
        if ($this->formula1 !== null && $this->type !== self::TYPE_NONE && $this->showErrorMessage === null) {
            $this->showErrorMessage();
        }

        return $this;
    }

    /**
     * Set formula 2 for data validation
     *
     * @param int|float|string|array|null $formula
     *
     * @return $this
     */
    public function setFormula2($formula): DataValidation
    {
        $this->formula2 = ($formula !== null) ? $this->checkFormula($formula) : null;

        return $this;
    }

    /**
     * Alias of setFormula1()
     *
     * @param int|float|string|array|null $formula
     *
     * @return $this
     */
    public function setFormula($formula): DataValidation
    {

        return $this->setFormula1($formula);
    }

    /**
     * @param string $operator
     * @param $formula1
     * @param $formula2
     *
     * @return $this
     */
    public function setOperator(string $operator, $formula1 = null, $formula2 = null): DataValidation
    {
        if (isset($this->aliases[$operator])) {
            $operator = $this->aliases[$operator];
        }
        if (!in_array($operator, $this->availableOperators)) {
            ExceptionDataValidation::throwNew('Invalid operator for data validation "' . $operator . '"');
        }
        if (is_array($formula1) && $formula2 === null) {
            if ($operator === self::OPERATOR_BETWEEN || $operator === self::OPERATOR_NOT_BETWEEN) {
                $formulas = array_values($formula1);
                $formula1 = $formulas[0] ?? null;
                $formula2 = $formulas[1] ?? null;
            }
            else {
                ExceptionDataValidation::throwNew('Formula 1 is array, scalar value expected');
            }
        }
        $this->operator = $operator;
        $this->setFormula1($formula1);
        $this->setFormula2($formula2);
        if ($this->type !== self::TYPE_NONE && $this->showErrorMessage === null) {
            $this->showErrorMessage();
        }

        return $this;
    }

    /**
     * @param Sheet $sheet
     * @param string $sqref
     *
     * @return $this
     */
    public function setSqref(Sheet $sheet, string $sqref): DataValidation
    {
        $this->sheet = $sheet;
        $this->sqref = $sqref;

        return $this;
    }

    /**
     * Allow blank value
     *
     * @param bool $allowBlank
     *
     * @return $this
     */
    public function allowBlank(?bool $allowBlank = true): DataValidation
    {
        $this->allowBlank = (int)$allowBlank;

        return $this;
    }

    /**
     * Show dropdown list
     *
     * @param bool $showDropDown
     *
     * @return $this
     */
    public function showDropDown(?bool $showDropDown = true): DataValidation
    {
        $this->showDropDown = (int)$showDropDown;

        return $this;
    }

    /**
     * Show input message
     *
     * @param bool|null $showInputMessage
     *
     * @return $this
     */
    public function showInputMessage(?bool $showInputMessage = true): DataValidation
    {
        $this->showInputMessage = (int)$showInputMessage;

        return $this;
    }

    /**
     * Error style (action in case of error)
     *
     * @param string $errorStyle
     *
     * @return $this
     */
    public function setErrorStyle(string $errorStyle): DataValidation
    {
        if (!in_array($errorStyle, $this->availableErrorStyles)) {
            ExceptionDataValidation::throwNew('Invalid error style for data validation "' . $errorStyle . '"');
        }
        $this->errorStyle = $errorStyle;

        return $this;
    }

    /**
     * Error message title
     *
     * @param string $errorTitle
     *
     * @return $this
     */
    public function setErrorTitle(string $errorTitle): DataValidation
    {
        $this->errorTitle = $errorTitle;

        return $this;
    }

    /**
     * Error message body
     *
     * @param string $error
     *
     * @return $this
     */
    public function setErrorMessage(string $error): DataValidation
    {
        $this->error = $error;

        return $this;
    }

    /**
     * Allow (or disallow) error message
     *
     * @param bool $showErrorMessage
     *
     * @return $this
     */
    public function showErrorMessage(?bool $showErrorMessage = true): DataValidation
    {
        $this->showErrorMessage = (int)$showErrorMessage;

        return $this;
    }

    /**
     * Set prompt
     *
     * @param string $promptMessage
     * @param string|null $promptTitle
     *
     * @return $this
     */
    public function setPrompt(string $promptMessage, ?string $promptTitle = null): DataValidation
    {
        $this->promptTitle = $promptTitle;
        $this->prompt = $promptMessage;
        if ($this->showInputMessage === null) {
            $this->showInputMessage = 1;
        }

        return $this;
    }

    /**
     * Set error message (title and body)
     *
     * @param string $errorMessage
     * @param string|null $errorTitle
     *
     * @return $this
     */
    public function setError(string $errorMessage, ?string $errorTitle = null): DataValidation
    {
        $this->error = $errorMessage;
        if ($errorTitle) {
            $this->errorTitle = $errorTitle;
        }
        if ($this->showErrorMessage === null) {
            $this->showErrorMessage = 1;
        }
        if ($this->errorStyle === null) {
            $this->errorStyle = self::STYLE_STOP;
        }

        return $this;
    }

    /**
     * @return array
     */
    public function getAttributes(): array
    {
        $attributes = [
            'type' => $this->type,
            'errorStyle' => $this->errorStyle,
            'operator' => $this->operator,
            'allowBlank' => $this->allowBlank,
            'showDropDown' => $this->showDropDown,
            'showInputMessage' => $this->showInputMessage,
            'showErrorMessage' => $this->showErrorMessage,
            'errorTitle' => $this->errorTitle,
            'error' => $this->error,
            'promptTitle' => $this->promptTitle,
            'prompt' => $this->prompt,
            'sqref' => $this->sqref,
        ];
        foreach ($attributes as $key => $value) {
            if ($value === null) {
                unset($attributes[$key]);
            }
        }

        return $attributes;
    }

    /**
     * @param $formulaConverter
     *
     * @return string
     */
    public function toXml($formulaConverter = null): string
    {
        $xml = '<dataValidation';
        foreach ($this->getAttributes() as $attribute => $value) {
            $xml .= ' ' . $attribute . '="' . $value . '"';
        }
        $xml .= '>';
        if ($this->formula1 !== null && $this->formula1 !== '') {
            if ($this->formula1[0] === '=') {
                $formula = ($formulaConverter ? $formulaConverter($this->formula1, $this->sqref) : substr($this->formula1, 1));
            }
            else {
                $formula = $this->formula1;
            }
            $xml .= '<formula1>' . $formula . '</formula1>';
        }
        if ($this->formula2 !== null && $this->formula2 !== '') {
            if ($this->formula2[0] === '=') {
                $formula = ($formulaConverter ? $formulaConverter($this->formula2, $this->sqref) : substr($this->formula2, 1));
            }
            else {
                $formula = $this->formula2;
            }
            $xml .= '<formula2>' . $formula . '</formula2>';
        }
        $xml .= '</dataValidation>';

        return $xml;
    }

}