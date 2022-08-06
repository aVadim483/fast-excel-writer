<?php

namespace avadim\FastExcelWriter;

use avadim\FastExcelWriter\Exception\Exception;

/**
 * Class Area
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
    protected $dimension = [];

    /** @var string */
    protected $range;

    /** @var int  */
    protected int $index = -1;

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
            $range = 'A' . ($sheet->rowCount + 1);
            $dimension = Excel::rangeDimensionRelative($range, $offset, true);
        }
        else {
            $dimension = Excel::rangeDimension($range, true);
        }
        if ($dimension['rowNum1'] <= $sheet->rowCount) {
            throw new Exception("Cannot make area range $range (row number must be greater then written rows)");
        }

        $this->sheet = $sheet;
        $this->dimension = $dimension;
        $coord = [
            ['row' => $this->dimension['rowNum1'], 'col' => $this->dimension['colNum1']],
            ['row' => $this->dimension['rowNum2'], 'col' => $this->dimension['colNum2']]
        ];
        $this->range = $this->dimension['range'];

        $this->setCoord($coord);
    }

    /**
     * @param array $coord
     *
     * @return $this
     */
    public function setCoord(array $coord)
    {
        $this->coord = [];
        foreach($coord as $addr) {
            if (isset($addr['row'], $addr['col'])) {
                $row = $addr['row'];
                $col = $addr['col'];
            } else {
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
    public function setIndex(int $index)
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
    public function setValue($cellAddress, $value, ?array $style = null)
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
    public function setFormula($cellAddress, $value, array $style = null)
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
    public function setStyle($cellAddress, array $style)
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
    public function addStyle($cellAddress, array $style)
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
    public function setFormat($cellAddress, string $format)
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->applayStyle($cellAddress, ['format' => $format], true);
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
    public function setColor($cellAddress, string $color)
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->applayStyle($cellAddress, ['color' => $color], true);
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
    public function setBackgroundColor($cellAddress, string $color)
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->applayStyle($cellAddress, ['fill' => $color], true);
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
    public function setFgColor($cellAddress, string $color)
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
    public function setBgColor($cellAddress, string $color)
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
    public function setOuterBorder($range, $style)
    {
        if (is_string($range) && $this->_validateAddressRange($range)) {
            $this->sheet->setOuterBorder($range, $style);
        }
        return $this;
    }

}

// EOF