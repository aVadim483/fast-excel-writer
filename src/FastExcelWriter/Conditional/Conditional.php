<?php

namespace avadim\FastExcelWriter\Conditional;

use avadim\FastExcelWriter\Exceptions\ExceptionConditionalFormatting;
use avadim\FastExcelWriter\Sheet;
use avadim\FastExcelWriter\StyleManager;

class Conditional
{
    /* Condition types */
    const CONDITION_CELL = 'cellIs';
    const CONDITION_TEXT = 'containsText';
    const CONDITION_EXPRESSION = 'expression';
    const CONDITION_COLOR_SCALE = 'colorScale';
    const CONDITION_DATA_BAR = 'dataBar';
    const CONDITION_ABOVE_AVERAGE = 'aboveAverage';
    const CONDITION_BELOW_AVERAGE = 'belowAverage';
    const CONDITION_UNIQUE_VALUES = 'uniqueValues';
    const CONDITION_DUPLICATE_VALUES = 'duplicateValues';
    const CONDITION_TOP10 = 'top10';

    /* Operator types */
    const OPERATOR_NONE = '';
    const OPERATOR_EQUAL = 'equal';
    const OPERATOR_NOT_EQUAL = 'notEqual';
    const OPERATOR_GREATER_THAN = 'greaterThan';
    const OPERATOR_GREATER_THAN_OR_EQUAL = 'greaterThanOrEqual';
    const OPERATOR_LESS_THAN = 'lessThan';
    const OPERATOR_LESS_THAN_OR_EQUAL = 'lessThanOrEqual';
    const OPERATOR_BETWEEN = 'between';
    const OPERATOR_NOT_BETWEEN = 'notBetween';
    const OPERATOR_CONTAINS = 'containsText';
    const OPERATOR_NOT_CONTAINS = 'notContains';
    const OPERATOR_BEGINS_WITH = 'beginsWith';
    const OPERATOR_ENDS_WITH = 'endsWith';

    /**
     * Condition type
     */
    protected string $conditionType;

    protected string $operator;

    protected array $style = [
        'font-color' => '#000000',
        'fill-color' => '#ffffff',
        'fill-pattern' => 'solid',
        'color-rgb' => [],
    ];

    protected Sheet $sheet;
    protected int $priority;
    protected string $sqref;

    protected string $dxfId;

    protected ?string $text = null;

    protected array $formula = [];

    protected array $cfvo = [];

    protected array $dataBarOptions = [
        'gradient' => true,
        'showValue' => true,
        'directionRtl' => null,
    ];
    protected array $topOptions = [
        'rank' => 0,
        'percent' => 0,
    ];

    protected static array $operatorTypes = [
        self::OPERATOR_NONE,
        self::OPERATOR_BEGINS_WITH,
        self::OPERATOR_ENDS_WITH,
        self::OPERATOR_EQUAL,
        self::OPERATOR_GREATER_THAN,
        self::OPERATOR_GREATER_THAN_OR_EQUAL,
        self::OPERATOR_LESS_THAN,
        self::OPERATOR_LESS_THAN_OR_EQUAL,
        self::OPERATOR_NOT_EQUAL,
        self::OPERATOR_CONTAINS,
        self::OPERATOR_NOT_CONTAINS,
        self::OPERATOR_BETWEEN,
        self::OPERATOR_NOT_BETWEEN,
    ];

    protected array $aliases = [
        '=' => self::OPERATOR_EQUAL,
        '!=' => self::OPERATOR_NOT_EQUAL,
        '<>' => self::OPERATOR_NOT_EQUAL,
        '>' => self::OPERATOR_GREATER_THAN,
        '>=' => self::OPERATOR_GREATER_THAN_OR_EQUAL,
        '<' => self::OPERATOR_LESS_THAN,
        '<=' => self::OPERATOR_LESS_THAN_OR_EQUAL,
        '!between' => self::OPERATOR_NOT_BETWEEN,
    ];

    protected static array $presetStyles = [
        'red' => ['font-color' => '#9C0006', 'fill-color' => '#FFC7CE', 'fill-pattern' => 'solid'],
        'yellow' => ['font-color' => '#9C5700', 'fill-color' => '#FFEB9C', 'fill-pattern' => 'solid'],
        'green' => ['font-color' => '#006100', 'fill-color' => '#C6EFCE', 'fill-pattern' => 'solid'],
    ];

    /**
     * Create a new Conditional
     */
    public function __construct(string $type, string $operator, $options, $style = null)
    {
        if (isset($this->aliases[$operator])) {
            $operator = $this->aliases[$operator];
        }
        if (!in_array($operator, self::$operatorTypes)) {
            ExceptionConditionalFormatting::throwNew('Invalid operator for conditional formatting "' . $operator . '"');
        }

        if (isset($options['formula'])) {
            $this->formula = (array)$options['formula'];
        }
        if (isset($options['cfvo'])) {
            $this->cfvo = $options['cfvo'];
        }
        if (isset($options['text'])) {
            $this->text = $options['text'];
        }

        if ($type === self::CONDITION_EXPRESSION) {
            if ($this->formula && $this->formula[0] !== '=') {
                $this->formula[0] = '=' . $this->formula[0];
            }
        }
        elseif ($type === self::CONDITION_TOP10 && isset($options['options'])) {
            $this->topOptions = ($options['options']);
        }

        if ($style) {
            $this->setStyle($style);
        }

        $this->conditionType = $type;
        $this->operator = $operator;
        if (($operator === self::OPERATOR_EQUAL || $operator === self::OPERATOR_NOT_EQUAL) && isset($this->formula[0]) && is_string($this->formula[0])) {
            if ($this->formula[0][0] !== '=') {
                $this->text = $this->formula[0];
                $this->formula[0] = '"' . $this->formula[0] . '"';
            }
        }

        foreach ($this->formula as $n => $formula) {
            $this->formula[$n] = ($formula === null) ? null : (string)$formula;
        }
    }

    /**
     * Cell value is compared to a specified value or formula
     *
     * @param string $operator
     * @param int|float|string|array $formula
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function make(string $operator, $formula, ?array $style = null): Conditional
    {
        return new self(self::CONDITION_CELL, $operator, ['formula' => $formula], $style);
    }

    /**
     * The cell value is equal to the given value
     *
     * @param int|float|string $value
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function equals($value, ?array $style = null): Conditional
    {
        return Conditional::make(self::OPERATOR_EQUAL, $value, $style);
    }

    /**
     * The cell value is not equal to the specified value
     *
     * @param int|float|string $value
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function notEquals($value, ?array $style = null): Conditional
    {
        return Conditional::make(self::OPERATOR_NOT_EQUAL, $value, $style);
    }

    /**
     * The cell value is greater than the specified value
     *
     * @param int|float|string $value
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function greaterThan($value, ?array $style = null): Conditional
    {
        return Conditional::make(self::OPERATOR_GREATER_THAN, $value, $style);
    }

    /**
     * The cell value is greater than or equal to the specified value
     *
     * @param int|float|string $value
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function greaterThanOrEqual($value, ?array $style = null): Conditional
    {
        return Conditional::make(self::OPERATOR_GREATER_THAN_OR_EQUAL, $value, $style);
    }

    /**
     * The cell value is less than the specified value
     *
     * @param int|float|string $value
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function lessThan($value, ?array $style = null): Conditional
    {
        return Conditional::make(self::OPERATOR_LESS_THAN, $value, $style);
    }

    /**
     * The cell value is less than or equal to the specified value
     *
     * @param int|float|string $value
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function lessThanOrEqual($value, ?array $style = null): Conditional
    {
        return Conditional::make(self::OPERATOR_LESS_THAN_OR_EQUAL, $value, $style);
    }

    /**
     * The cell value is between two given values
     *
     * @param int[]|float[] $values
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function between(array $values, ?array $style = null): Conditional
    {
        return Conditional::make(self::OPERATOR_BETWEEN, $values, $style);
    }

    /**
     * The cell value is between two given values
     *
     * @param int[]|float[] $values
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function notBetween(array $values, ?array $style = null): Conditional
    {
        return Conditional::make(self::OPERATOR_NOT_BETWEEN, $values, $style);
    }

    /**
     * Applies a style if the cell value contains the specified text.
     *
     * @param string $text
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function contains(string $text, ?array $style = null): Conditional
    {
        return new self(self::CONDITION_TEXT, self::OPERATOR_CONTAINS, ['text' => $text], $style);
    }

    /**
     * Applies a style if the cell value does not contain the specified text.
     *
     * @param string $text
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function notContains(string $text, ?array $style = null): Conditional
    {
        return new self(self::CONDITION_TEXT, self::OPERATOR_NOT_CONTAINS, ['text' => $text], $style);
    }

    /**
     * Applies a style if the cell value starts with the specified text
     *
     * @param string $text
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function beginsWith(string $text, ?array $style = null): Conditional
    {
        return new self(self::CONDITION_TEXT, self::OPERATOR_BEGINS_WITH, ['text' => $text], $style);
    }

    /**
     * Applies a style if the cell value ends with the specified text
     *
     * @param string $text
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function endsWith(string $text, ?array $style = null): Conditional
    {
        return new self(self::CONDITION_TEXT, self::OPERATOR_ENDS_WITH, ['text' => $text], $style);
    }

    /**
     * Applies the style if the expression evaluates to TRUE
     *
     * @param string $formula
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function expression(string $formula, ?array $style = null): Conditional
    {
        return new self(self::CONDITION_EXPRESSION, '', ['formula' => $formula], $style);
    }

    /**
     * Applies a style if the cell is empty
     *
     * @param string|null $cell
     * @param array|null $style
     *
     * @return Conditional
     */
    public static function isEmpty(?string $cell = null, ?array $style = null): Conditional
    {
        if ($cell === null) {
            return self::expression('=ISBLANK(RC)', $style);
        }
        return self::expression('=ISBLANK(' . strtoupper($cell) . ')', $style);
    }

    /**
     * @param string $color1
     * @param string $color2
     * @param string|null $color3
     *
     * @return Conditional
     */
    public static function colorScale(string $color1, string $color2, ?string $color3 = null): Conditional
    {
        $cfvo = [
            ['type' => 'min'],
            ['type' => 'max'],
        ];
        if ($color3) {
            $cfvo[] = ['type' => 'percentile', 'val' => 50];
        }
        $style = ['color-rgb' => [$color1, $color2, $color3]];

        return new self(self::CONDITION_COLOR_SCALE, '', ['cfvo' => $cfvo], $style);
    }

    /**
     * @param string $color
     *
     * @return Conditional
     */
    public static function colorScaleMax(string $color): Conditional
    {
        $cfvo = [
            ['type' => 'min'],
            ['type' => 'max'],
        ];
        $style = ['color-rgb' => ['#ffffff', $color]];

        return new self(self::CONDITION_COLOR_SCALE, '', ['cfvo' => $cfvo], $style);
    }

    /**
     * @param string $color
     *
     * @return Conditional
     */
    public static function colorScaleMin(string $color): Conditional
    {
        $cfvo = [
            ['type' => 'min'],
            ['type' => 'max'],
        ];
        $style = ['color-rgb' => [$color, '#ffffff']];

        return new self(self::CONDITION_COLOR_SCALE, '', ['cfvo' => $cfvo], $style);
    }

    /**
     * @param array $values
     * @param string $color1
     * @param string $color2
     * @param string|null $color3
     *
     * @return Conditional
     */
    public static function colorScaleNum(array $values, string $color1, string $color2, ?string $color3 = null): Conditional
    {
        $cfvo = [];
        foreach ($values as $val) {
            $cfvo[] = ['type' => 'num', 'val' => $val];
        }
        $style = ['color-rgb' => [$color1, $color2, $color3]];

        return new self(self::CONDITION_COLOR_SCALE, '', ['cfvo' => $cfvo], $style);
    }

    /**
     * Colored data bar inside a cell
     *
     * @param string $color
     *
     * @return Conditional
     */
    public static function dataBar(string $color): Conditional
    {
        $cfvo = [
            ['type' => 'min', 'val' => null], // 0
            ['type' => 'max', 'val' => null], // 100
        ];
        $style = ['color-rgb' => [$color]];

        return new self(self::CONDITION_DATA_BAR, '', ['cfvo' => $cfvo], $style);
    }

    /**
     * Enables or disables the gradient style of the bars
     *
     * @param bool $value
     *
     * @return $this
     */
    public function setGradient(bool $value): Conditional
    {
        $this->dataBarOptions['gradient'] = $value;

        return $this;
    }

    /**
     * Controls the display of the value in a cell
     *
     * @param bool $value
     *
     * @return $this
     */
    public function setShowValue(bool $value): Conditional
    {
        $this->dataBarOptions['showValue'] = $value;

        return $this;
    }

    /**
     * Determines the direction of the bars
     *
     * @param bool $value
     *
     * @return $this
     */
    public function setDirectionRtl(bool $value): Conditional
    {
        $this->dataBarOptions['directionRtl'] = $value;

        return $this;
    }

    /**
     * @param array $style
     *
     * @return Conditional
     */
    public static function aboveAverage(array $style): Conditional
    {
        return new self(self::CONDITION_ABOVE_AVERAGE, '', null, $style);
    }

    /**
     * @param array $style
     *
     * @return Conditional
     */
    public static function belowAverage(array $style): Conditional
    {
        $options = [
            'aboveAverage' => 0,
        ];

        return new self(self::CONDITION_BELOW_AVERAGE, '', $options, $style);
    }

    /**
     * @param array $style
     *
     * @return Conditional
     */
    public static function uniqueValues(array $style): Conditional
    {
        return new self(self::CONDITION_UNIQUE_VALUES, '', null, $style);
    }

    /**
     * @param array $style
     *
     * @return Conditional
     */
    public static function duplicateValues(array $style): Conditional
    {
        return new self(self::CONDITION_DUPLICATE_VALUES, '', null, $style);
    }

    /**
     * @param int $rank
     * @param array $style
     *
     * @return Conditional
     */
    public static function top(int $rank, array $style): Conditional
    {
        $options = [
            'rank' => $rank,
            'percent' => 0,
        ];

        return new self(self::CONDITION_TOP10, '', ['options' => $options], $style);
    }

    /**
     * @param int $rank
     * @param array $style
     *
     * @return Conditional
     */
    public static function topPercent(int $rank, array $style): Conditional
    {
        $options = [
            'rank' => $rank,
            'percent' => 1,
        ];

        return new self(self::CONDITION_TOP10, '', ['options' => $options], $style);
    }

    /**
     * @param int $rank
     * @param array $style
     *
     * @return Conditional
     */
    public static function low(int $rank, array $style): Conditional
    {
        $options = [
            'rank' => $rank,
            'percent' => 0,
            'bottom' => 1,
        ];

        return new self(self::CONDITION_TOP10, '', ['options' => $options], $style);
    }

    /**
     * @param int $rank
     * @param array $style
     *
     * @return Conditional
     */
    public static function lowPercent(int $rank, array $style): Conditional
    {
        $options = [
            'rank' => $rank,
            'percent' => 1,
            'bottom' => 1,
        ];

        return new self(self::CONDITION_TOP10, '', ['options' => $options], $style);
    }

    /**
     * @param string|array $style
     *
     * @return $this
     */
    public function setStyle($style): Conditional
    {
        if (is_string($style)) {
            $style = strtolower($style);
            if (!isset(self::$presetStyles[$style])) {
                ExceptionConditionalFormatting::throwNew('Invalid style name for conditional formatting');
            }
            $this->style = self::$presetStyles[$style];
        }
        else {
            $this->style = $style;
        }

        return $this;
    }

    public function setFontColor($color): Conditional
    {
        $this->style['font-color'] = $color;

        return $this;
    }

    public function setFillColor($color): Conditional
    {
        $this->style['fill-color'] = $color;

        return $this;
    }

    /**
     * @param Sheet $sheet
     * @param string $sqref
     *
     * @return $this
     */
    public function setSqref(Sheet $sheet, string $sqref): Conditional
    {
        $this->sheet = $sheet;
        $this->sqref = $sqref;
        $dxfId = $sheet->excel->addStyleDxfs($this->getStyle());
        $this->setDxfId($dxfId);
        $this->priority = count($sheet->getConditionalFormatting()) + 1;
        if (isset($this->dataBarOptions['directionRtl']) && $sheet->isRightToLeft()) {
            $this->dataBarOptions['directionRtl'] = true;
        }

        return $this;
    }

    /**
     * @param int $dxfId
     *
     * @return $this
     */
    public function setDxfId(int $dxfId): Conditional
    {
        $this->dxfId = $dxfId;

        return $this;
    }

    /**
     * @return array
     */
    public function getStyle(): ?array
    {
        return $this->style;
    }

    /**
     * @param array $attributes
     *
     * @return string
     */
    protected function _attr(array $attributes): string
    {
        $result = '';
        foreach ($attributes as $attribute => $value) {
            if ($value !== null) {
                $result .= ' ' . $attribute . '="' . $value . '"';
            }
        }

        return $result;
    }

    /**
     * @param int $priority
     * @param $formulaConverter
     *
     * @return string
     */
    public function toXml(int $priority, $formulaConverter = null): string
    {
        $xml = '<conditionalFormatting sqref="' . $this->sqref . '">';
        $firstCell = strpos($this->sqref, ':') ? strstr($this->sqref, ':', true) : $this->sqref;
        if ($this->conditionType === self::CONDITION_TEXT) {
            if ($this->operator === self::OPERATOR_NOT_CONTAINS) {
                $formula = 'ISERROR(SEARCH("' . $this->text . '",' . $firstCell . '))';
            }
            elseif ($this->operator === self::OPERATOR_BEGINS_WITH) {
                $formula = 'LEFT(' . $firstCell . ',' . mb_strlen($this->text) . ')="' . $this->text . '"';
            }
            elseif ($this->operator === self::OPERATOR_ENDS_WITH) {
                $formula = 'RIGHT(' . $firstCell . ',' . mb_strlen($this->text) . ')="' . $this->text . '"';
            }
            else {
                $formula = 'NOT(ISERROR(SEARCH("' . $this->text . '",' . $firstCell . ')))';
            }
            $xml .= '<cfRule type="' . $this->conditionType . '" dxfId="' . $this->dxfId . '" priority="' . $priority . '" operator="' . $this->operator . '" text="' . $this->text . '">';
            $xml .= '<formula>' . $formula . '</formula>';
            $xml .= '</cfRule>';
        }
        elseif ($this->conditionType === self::CONDITION_COLOR_SCALE || $this->conditionType === self::CONDITION_DATA_BAR) {
            $xml .= '<cfRule type="' . $this->conditionType . '" priority="' . $priority . '">';
            $xml .= '<' . $this->conditionType . '>';
            foreach ($this->cfvo as $item) {
                if (isset($item['val'])) {
                    $xml .= '<cfvo type="' . $item['type'] . '" val="' . $item['val'] . '"/>';
                }
                else {
                    $xml .= '<cfvo type="' . $item['type'] . '"/>';
                }
            }
            foreach ($this->style['color-rgb'] as $color) {
                if ($color) {
                    $xml .= '<color rgb="' . StyleManager::normalizeColor($color) . '"/>';
                }
            }

            if ($this->conditionType === self::CONDITION_DATA_BAR) {
                $xml .= '<gradient val="' . (!empty($this->dataBarOptions['gradient']) ? 'true' : 'false') . '"/>';
                $xml .= '<showValue val="' . (!empty($this->dataBarOptions['showValue']) ? 'true' : 'false') . '"/>';
                if (!empty($this->dataBarOptions['directionRtl'])) {
                    $xml .= '<direction rtl="true"/>';
                }
            }

            $xml .= '</' . $this->conditionType . '>';
            $xml .= '</cfRule>';
        }
        elseif ($this->conditionType === self::CONDITION_TOP10) {
            $attributes = [
                'type' => $this->conditionType,
                'dxfId' => $this->dxfId,
                'priority' => $priority,
                'rank' => $this->topOptions['rank'],
                'percent' => ($this->topOptions['percent'] ? 1 : 0),
            ];
            if (!empty($this->topOptions['bottom'])) {
                $attributes['bottom'] = $this->topOptions['bottom'];
            }
            $xml .= '<cfRule' . $this->_attr($attributes) . '/>';
        }
        else {
            if ($this->conditionType === self::CONDITION_BELOW_AVERAGE) {
                $type = self::CONDITION_ABOVE_AVERAGE;
                $aboveAverage = 0;
            }
            else {
                $type = $this->conditionType;
                $aboveAverage = null;
            }
            $attributes = [
                'type' => $type,
                'dxfId' => $this->dxfId,
                'priority' => $priority,
                'operator' => $this->operator ?: null,
                'text' => $this->text ?: null,
                'aboveAverage' => $aboveAverage,
            ];
            $xml .= '<cfRule' . $this->_attr($attributes) . '>';

            foreach ($this->formula as $formula) {
                if ($formula !== null && $formula !== '') {
                    if ($formula[0] === '=') {
                        $formula = ($formulaConverter ? $formulaConverter($formula, $firstCell) : substr($formula, 1));
                    }
                    $xml .= '<formula>' . $formula . '</formula>';
                }
            }
            $xml .= '</cfRule>';
        }
        $xml .= '</conditionalFormatting>';

        return $xml;
    }

}