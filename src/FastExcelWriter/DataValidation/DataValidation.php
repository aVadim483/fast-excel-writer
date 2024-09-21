<?php

namespace avadim\FastExcelWriter\DataValidation;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Exceptions\ExceptionDataValidation;

class DataValidation
{
    /* Data validation types */
    const TYPE_CUSTOM = 'custom'; // Data validation which uses a custom formula to check the cell value
    const TYPE_DATE = 'date'; // Data validation which checks for date values satisfying the given condition
    const TYPE_DECIMAL = 'decimal'; // Data validation which checks for decimal values satisfying the given condition
    const TYPE_LIST = 'list'; // Data validation which checks for a value matching one of list of values
    const TYPE_NONE = 'none'; // No data validation. Why do we need this type?
    const TYPE_TEXT_LENGTH = 'textLength'; // Data validation which checks for text values, whose length satisfies the given condition
    const TYPE_TEXTLENGTH = 'textLength';
    const TYPE_TIME = 'time'; // Data validation which checks for time values satisfying the given condition
    const TYPE_WHOLE = 'whole'; // Data validation which checks for whole number values satisfying the given condition


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

    public function __construct($type)
    {
        $this->type = $type;
    }

    /**
     * @param $type
     *
     * @return DataValidation
     */
    public static function make($type): DataValidation
    {
        return new self($type);
    }

    /**
     * @param string $operator
     * @param string|int|array $values
     *
     * @return DataValidation
     */
    public static function integer(string $operator, $values): DataValidation
    {
        $validation = new self(self::TYPE_WHOLE);
        if (is_array($values)) {
            $values = array_values($values);
            $validation->setOperator($operator, $values[0] ?? null, $values[0] ?? null);
        }
        else {
            $validation->setOperator($operator, $values);
        }

        return $validation;
    }

    /**
     * @param string $operator
     * @param string|int|array $values
     *
     * @return DataValidation
     */
    public static function decimal(string $operator, $values): DataValidation
    {
        $validation = new self(self::TYPE_DECIMAL);
        if (is_array($values)) {
            $values = array_values($values);
            $validation->setOperator($operator, $values[0] ?? null, $values[0] ?? null);
        }
        else {
            $validation->setOperator($operator, $values);
        }

        return $validation;
    }

    /**
     * @param string $operator
     * @param string|int|array $values
     *
     * @return DataValidation
     */
    public static function date(string $operator, $values): DataValidation
    {
        $validation = new self(self::TYPE_DATE);
        if (is_array($values)) {
            $values = array_values($values);
            $validation->setOperator($operator, $values[0] ?? null, $values[0] ?? null);
        }
        else {
            $validation->setOperator($operator, $values);
        }

        return $validation;
    }

    /**
     * @param array $values
     *
     * @return DataValidation
     */
    public static function dropDown(array $values): DataValidation
    {
        $validation = new self(self::TYPE_LIST);
        $validation->setFormula1($values);

        return $validation;
    }

    /**
     * @param $formula
     *
     * @return string
     */
    protected function checkFormula($formula): string
    {
        if (is_string($formula) && $formula && $formula[0] === '=') {
            $dimension = Excel::rangeDimension(substr($formula, 1));
            $result = $dimension['absAddress'];
        }
        elseif (is_array($formula)) {
            $result = '"' . implode(',', $formula) . '"';
        }
        elseif (is_float($formula)) {
            $result = str_replace(',' , '.', (string)$formula);
        }
        else {
            $result = (string)$formula;
        }

        return $result;
    }

    /**
     * @param int|float|string|array|null $formula
     *
     * @return $this
     */
    public function setFormula1($formula): DataValidation
    {
        $this->formula1 = ($formula !== null) ? $this->checkFormula($formula) : null;

        return $this;
    }


    public function setFormula2($formula): DataValidation
    {
        $this->formula2 = ($formula !== null) ? $this->checkFormula($formula) : null;

        return $this;
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
        $this->operator = $operator;
        $this->setFormula1($formula1);
        $this->setFormula2($formula2);

        return $this;
    }

    /**
     * @param string $sqref
     *
     * @return $this
     */
    public function setSqref(string $sqref): DataValidation
    {
        $this->sqref = $sqref;

        return $this;
    }

    /**
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
     * @param bool $showDropDown
     *
     * @return $this
     */
    public function showDropDown(?bool $showDropDown = true): DataValidation
    {
        $this->showDropDown = (int)$showDropDown;

        return $this;
    }


    public function showInputMessage(?bool $showInputMessage = true): DataValidation
    {
        $this->showInputMessage = (int)$showInputMessage;

        return $this;
    }

    /**
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
     * @param bool $showErrorMessage
     *
     * @return $this
     */
    public function showErrorMessage(bool $showErrorMessage): DataValidation
    {
        $this->showErrorMessage = (int)$showErrorMessage;

        return $this;
    }

    /**
     * @param $formula1
     * @param $formula2
     *
     * @return DataValidation
     */
    public function setOperatorBetween($formula1, $formula2): DataValidation
    {
        return $this->setOperator(self::OPERATOR_BETWEEN)
            ->setFormula1($formula1)
            ->setFormula2($formula2);
    }

    /**
     * @param $formula1
     * @param $formula2
     *
     * @return DataValidation
     */
    public function setOperatorNotBetween($formula1, $formula2): DataValidation
    {
        return $this->setOperator(self::OPERATOR_NOT_BETWEEN)->setFormula1($formula1)->setFormula2($formula2);
    }

    /**
     * @param $formula
     *
     * @return DataValidation
     */
    public function setOperatorEqual($formula): DataValidation
    {
        return $this->setOperator(self::OPERATOR_EQUAL)->setFormula1($formula)->setFormula2(null);
    }

    /**
     * @param $formula
     *
     * @return DataValidation
     */
    public function setOperatorNotEqual($formula): DataValidation
    {
        return $this->setOperator(self::OPERATOR_NOT_EQUAL)->setFormula1($formula)->setFormula2(null);
    }

    /**
     * @param $formula
     *
     * @return DataValidation
     */
    public function setOperatorGreaterThan($formula): DataValidation
    {
        return $this->setOperator(self::OPERATOR_GREATER_THAN)->setFormula1($formula)->setFormula2(null);
    }

    /**
     * @param $formula
     *
     * @return DataValidation
     */
    public function setOperatorGreaterThanOrEqual($formula): DataValidation
    {
        return $this->setOperator(self::OPERATOR_GREATER_THAN_OR_EQUAL)->setFormula1($formula)->setFormula2(null);
    }

    /**
     * @param $formula
     *
     * @return DataValidation
     */
    public function setOperatorLessThan($formula): DataValidation
    {
        return $this->setOperator(self::OPERATOR_LESS_THAN)->setFormula1($formula)->setFormula2(null);
    }

    /**
     * @param $formula
     *
     * @return DataValidation
     */
    public function setOperatorLessThanOeEqual($formula): DataValidation
    {
        return $this->setOperator(self::OPERATOR_LESS_THAN_OR_EQUAL)->setFormula1($formula)->setFormula2(null);
    }

    /**
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
     * @return string
     */
    public function toXml(): string
    {
        $xml = '<dataValidation';
        foreach ($this->getAttributes() as $attribute => $value) {
            $xml .= ' ' . $attribute . '="' . $value . '"';
        }
        $xml .= '>';
        if ($this->formula1 !== null) {
            $xml .= '<formula1>' . $this->formula1 . '</formula1>';
        }
        if ($this->formula2 !== null) {
            $xml .= '<formula2>' . $this->formula2 . '</formula2>';
        }
        $xml .= '</dataValidation>';

        return $xml;
    }

}