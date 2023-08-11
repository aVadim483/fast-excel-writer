<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;

/**
 * Class Area
 *
 * @method Area applyRowHeight(float $height)
 * @method Area applyBorder(string $style, ?string $color = '#000000')
 * @method Area applyBorderLeft(string $style, ?string $color = '#000000')
 * @method Area applyBorderRight(string $style, ?string $color = '#000000')
 * @method Area applyBorderTop(string $style, ?string $color = '#000000')
 * @method Area applyBorderBottom(string $style, ?string $color = '#000000')
 * @method Area applyOuterBorder(string $style, ?string $color = '#000000')
 * @method Area applyInnerBorder(string $style, ?string $color = '#000000')
 * @method Area applyFont(string $fontName, ?int $fontSize = null, ?string $fontStyle = null, ?string $fontColor = null)
 * @method Area applyFontName(string $fontName)
 * @method Area applyFontSize(float $fontSize)
 * @method Area applyFontStyle(string $fontStyle)
 * @method Area applyFontStyleBold()
 * @method Area applyFontStyleItalic()
 * @method Area applyFontStyleUnderline(?bool $double = false)
 * @method Area applyFontStyleStrikethrough()
 * @method Area applyFontColor(string $fontColor)
 * @method Area applyColor(string $color)
 * @method Area applyFillColor(string $color)
 * @method Area applyBgColor(string $color)
 * @method Area applyTextAlign(string $textAlign, ?string $verticalAlign = null)
 * @method Area applyVerticalAlign(string $verticalAlign)
 * @method Area applyTextCenter()
 * @method Area applyTextWrap(bool $textWrap)
 * @method Area applyTextColor(string $color)
 *
 * @package avadim\FastExcelWriter
 */
class Area
{
    /** @var Sheet */
    protected Sheet $sheet;

    /** @var array[]  */
    protected array $coord;

    /** @var array  */
    protected array $dimension = [];

    /** @var string */
    protected $range;

    /** @var int  */
    protected int $index = -1;

    protected int $currentColNum;

    protected int $currentRowNum;

    /**
     * Area constructor
     *
     * @param Sheet $sheet
     * @param string|array $range
     */
    public function __construct(Sheet $sheet, $range)
    {
        if (is_string($range) && preg_match('/^(-)?R(\d+)(-)?C(\d+)/i', $range)) {
            $offset = $range;
            $range = 'A' . ($sheet->rowCountWritten + 1);
            $dimension = Excel::rangeDimensionRelative($range, $offset, true);
        }
        else {
            $dimension = Excel::rangeDimension($range, true);
        }
        if ($dimension['rowNum1'] <= $sheet->rowCountWritten) {
            throw new Exception("Cannot make area range $range (row number must be greater then written rows)");
        }

        $this->sheet = $sheet;
        $this->dimension = $dimension;
        $coord = [
            ['row' => $this->dimension['rowNum1'], 'col' => $this->dimension['colNum1']],
            ['row' => $this->dimension['rowNum2'], 'col' => $this->dimension['colNum2']]
        ];
        $this->range = $this->dimension['range'];
        $this->currentRowNum = $this->dimension['rowNum1'];
        $this->currentColNum = $this->dimension['colNum1'];

        $this->setCoord($coord);
        $this->moveTo($dimension['cell1']);
    }

    /**
     * @param array $coord
     *
     * @return $this
     */
    public function setCoord(array $coord): Area
    {
        $this->coord = [];
        foreach($coord as $addr) {
            if (isset($addr['row'], $addr['col'])) {
                $row = $addr['row'];
                $col = $addr['col'];
            }
            else {
                [$row, $col] = $addr;
            }
            $this->coord[] = ['row' => $row, 'col' => $col];
        }
        return $this;
    }

    /**
     * Set index of area in sheet
     *
     * @param int $index
     *
     * @return $this
     */
    public function setIndex(int $index): Area
    {
        $this->index = $index;

        return $this;
    }

    /**
     * Get index of area in sheet
     *
     * @return int
     */
    public function getIndex(): int
    {
        return $this->index;
    }

    /**
     * @return array[]
     */
    public function getCoord(): array
    {
        return $this->coord;
    }

    /**
     * @return string
     */
    public function getBeginAddress(): string
    {
        $coord = $this->getCoord();

        return Excel::cellAddress($coord[0]['row'], $coord[0]['col']);
    }

    /**
     * @return string
     */
    public function getEndAddress(): string
    {
        $coord = $this->getCoord();

        return Excel::cellAddress($coord[1]['row'], $coord[1]['col']);
    }

    /**
     * @return string
     */
    public function getOffsetAddress($offset): string
    {
        $coord = $this->getCoord();

        return Excel::cellAddress($coord[1]['row'], $coord[1]['col']);
    }

    /**
     * @param string|array $cellAddress
     * @param array|null $index
     * @param array|null $offset
     *
     * @return bool
     */
    protected function _validateAddressRange(&$cellAddress, array &$index = null, array &$offset = null): bool
    {
        if ($cellAddress) {
            if (is_string($cellAddress)) {
                $offset = null;
                $cellAddress = strtoupper($cellAddress);
                if (strpos($cellAddress, 'R') === 0 && strpos($cellAddress, 'C')) {
                    $offset = Excel::rangeRelOffsets($cellAddress);
                }
                if ($offset && isset($offset[0], $offset[1], $offset[2], $offset[3])) {
                    $dim = Excel::rangeDimensionRelative($this->dimension['cell1'], $offset);
                    if (strpos($cellAddress, ':')) {
                        $cellAddress = $dim['cell1'] . ':' . $dim['cell2'];
                    }
                    else {
                        $cellAddress = $dim['cell1'];
                    }
                    return true;
                }
                $dimension = Excel::rangeDimension($cellAddress);
                if ($dimension) {
                    $idxAddress = ['row' => $dimension['row'], 'col' => $dimension['col']];
                }
                else {
                    $idxAddress = ['row' => -1, 'col' => -1];
                }

            }
            else {
                // $cellAddress is array
                $idxAddress = $cellAddress;
            }
            if ($idxAddress['row'] < $this->coord[0]['row'] || $idxAddress['row'] > $this->coord[1]['row']
                || $idxAddress['col'] < $this->coord[0]['col'] || $idxAddress['col'] > $this->coord[1]['col']) {
                throw new Exception('Address "' . $cellAddress . '" is outside area "' . $this->range . '"');
            }
            $index = $idxAddress;
            return true;
        }
        return false;
    }

    /**
     * Write value to cell
     *
     * setValue('A2', $value)
     * setValue(['col' => 3, 'row' => 1], $value) - equals to 'C1'
     * setValue('A2:C2', $value) - merge cells and write value
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param array|null $style
     *
     * @return $this
     */
    public function setValue($cellAddress, $value, ?array $style = null): Area
    {
        if ($this->_validateAddressRange($cellAddress)) {
            $this->sheet->setValue($cellAddress, $value, $style);
        }

        return $this;
    }

    /**
     * Write formula to cell
     *
     * @param string|array $cellAddress
     * @param mixed $value
     * @param array|null $style
     *
     * @return $this
     */
    public function setFormula($cellAddress, $value, array $style = null): Area
    {
        if ($this->_validateAddressRange($cellAddress)) {
            $this->sheet->setFormula($cellAddress, $value, $style);
        }

        return $this;
    }

    /**
     * Set style (old styles wil be replaced)
     *
     * @param string|array $cellAddress
     * @param array $style
     *
     * @return $this
     */
    public function setStyle($cellAddress, array $style): Area
    {
        if ($this->_validateAddressRange($cellAddress)) {
            $this->sheet->setStyle($cellAddress, $style);
        }
        return $this;
    }

    /**
     * New style wil be added to old style (if they exists)
     *
     * @param string|array $cellAddress
     * @param array $style
     *
     * @return $this
     */
    public function addStyle($cellAddress, array $style): Area
    {
        if ($this->_validateAddressRange($cellAddress)) {
            $this->sheet->addStyle($cellAddress, $style);
        }
        return $this;
    }

    /**
     * Set format of values (old styles wil be replaced)
     *
     * @param string|array $cellAddress
     * @param string $format
     *
     * @return $this
     */
    public function setFormat($cellAddress, string $format): Area
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->setStyle($cellAddress, ['format' => $format], true);
        }
        return $this;
    }

    /**
     * Set text color
     *
     * @param string|array $cellAddress
     * @param string $color
     *
     * @return $this
     */
    public function setColor($cellAddress, string $color): Area
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->setStyle($cellAddress, ['font-color' => $color], true);
        }
        return $this;
    }

    /**
     * Set background color
     *
     * @param string|array $cellAddress
     * @param string $color
     *
     * @return $this
     */
    public function setBackgroundColor($cellAddress, string $color): Area
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->setStyle($cellAddress, ['fill-color' => $color], true);
        }
        return $this;
    }

    /**
     * Set text color, alias of setColor()
     *
     * @param string|array $cellAddress
     * @param string $color
     *
     * @return $this
     */
    public function setFgColor($cellAddress, string $color): Area
    {
        return $this->setColor($cellAddress, $color);
    }

    /**
     * Set background color, alias of setBackgroundColor()
     *
     * @param string|array $cellAddress
     * @param string $color
     *
     * @return $this
     */
    public function setBgColor($cellAddress, string $color): Area
    {
        return $this->setBackgroundColor($cellAddress, $color);
    }

    /**
     * Set outer border
     *
     * @param string|array $range
     * @param string|array $style
     *
     * @return $this
     */
    public function setOuterBorder($range, $style): Area
    {
        if (is_string($range) && $this->_validateAddressRange($range)) {
            $this->sheet->setOuterBorder($range, $style);
        }
        return $this;
    }

    /**
     * @param array|string $range
     *
     * @return $this
     */
    public function withRange($range): Area
    {
        if ($this->_validateAddressRange($range)) {
            $this->sheet->withRange($range);
        }

        return $this;
    }

    /**
     * Move the cursor to the specified address
     *
     * @param $cellAddress
     *
     * @return $this
     */
    public function moveTo($cellAddress): Area
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress, $numAddress)) {
            $this->currentColNum = $numAddress['col'];
            $this->currentRowNum = $numAddress['row'];
        }

        return $this;
    }

    /**
     * Write value to the cursor position and move cursor to the next cell in the row
     *
     * @return $this
     */
    public function nextCell(): Area
    {
        $this->currentColNum++;

        return $this;
    }

    /**
     * Move the cursor to the next cell in the row
     *
     * @return $this
     */
    public function nextRow(): Area
    {
        $this->currentRowNum++;
        $this->currentColNum = $this->dimension['colNum1'];

        return $this;
    }

    /**
     * Write a value to the cursor position and move the cursor to the next cell in the row
     *
     * @return $this
     */
    public function writeCell($value): Area
    {
        $this->setValue(['col' => $this->currentColNum++, 'row' => $this->currentRowNum], $value);

        return $this;
    }

    /**
     * Write a row data to the current row and move the cursor to the next row
     *
     * @param array $row
     *
     * @return $this
     */
    public function writeRow(array $row): Area
    {
        foreach ($row as $cell) {
            $this->writeCell($cell);
        }
        $this->nextRow();

        return $this;
    }

    /**
     * @param string $name
     * @param array $arguments
     *
     * @return $this
     *
     * @throws \Exception
     */
    public function __call(string $name, array $arguments)
    {
        if (strpos($name, 'apply') === 0 && method_exists($this->sheet, $name)) {
            call_user_func_array([$this->sheet, $name], $arguments);

            return $this;
        }

        $trace = debug_backtrace();
        $error = 'Uncaught Error: Call to undefined method ' . get_class() . '::' . $name
            . ' (called in ' . $trace[0]['file'] . ' on line ' . $trace[0]['line'] . ')';
        throw new \Exception($error);
    }
}

// EOF