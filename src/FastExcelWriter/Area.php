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
    protected $sheet;

    /** @var array[]  */
    protected $coord;

    /** @var array  */
    protected $dimension = [];

    /** @var string */
    protected $range;

    /** @var int  */
    protected $index = -1;

    /**
     * Area constructor
     *
     * @param Sheet $sheet
     * @param string $range
     */
    public function __construct($sheet, $range)
    {
        if (is_string($range) && preg_match('/^(-)?R(\d+)(-)?C(\d+)/i', $range)) {
            $offset = $range;
            $range = 'A' . ($sheet->rowCount + 1);
            $dimension = Excel::rangeDimensionRelative($range, $offset, true);
        } else {
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
     * @param $coord
     *
     * @return $this
     */
    public function setCoord($coord)
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
     * Index of area in sheet
     *
     * @return int
     */
    public function getIndex()
    {
        return $this->index;
    }

    /**
     * @return array[]
     */
    public function getCoord()
    {
        return $this->coord;
    }

    /**
     * @param string|array $cellAddress
     * @param array|null $index
     * @param array|null $offset
     *
     * @return bool
     */
    protected function _validateAddressRange(&$cellAddress, array &$index = null, array &$offset = null)
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
                    } else {
                        $cellAddress = $dim['cell1'];
                    }
                    return true;
                }
                $idxAddress = Excel::rangeRowColNumbers($cellAddress);
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
     * Set style (old styles wil be replaced)
     *
     * @param string|array $cellAddress
     * @param $format
     *
     * @return $this
     */
    public function setFormat($cellAddress, $format)
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->applayStyle($cellAddress, ['format' => $format], true);
        }
        return $this;
    }

    /**
     * @param string|array $cellAddress
     * @param $color
     *
     * @return $this
     */
    public function setColor($cellAddress, $color)
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->applayStyle($cellAddress, ['color' => $color], true);
        }
        return $this;
    }

    /**
     * @param string|array $cellAddress
     * @param $color
     *
     * @return $this
     */
    public function setBakgroundColor($cellAddress, $color)
    {
        if (is_string($cellAddress) && $this->_validateAddressRange($cellAddress)) {
            $this->sheet->applayStyle($cellAddress, ['fill' => $color], true);
        }
        return $this;
    }

    /**
     * @param string|array $cellAddress
     * @param $color
     *
     * @return $this
     */
    public function setFgColor($cellAddress, $color)
    {
        return $this->setColor($cellAddress, $color);
    }

    /**
     * @param string|array $cellAddress
     * @param $color
     *
     * @return $this
     */
    public function setBgColor($cellAddress, $color)
    {
        return $this->setBakgroundColor($cellAddress, $color);
    }

    /**
     * @param $range
     * @param $style
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