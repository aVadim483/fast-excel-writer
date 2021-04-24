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
        if ($dimension['rowIndex1'] <= $sheet->rowCount) {
            throw new Exception("Cannot make area range $range (row number must be greater then written rows)");
        }

        $this->sheet = $sheet;
        $this->dimension = $dimension;
        $coord = [
            ['row' => $this->dimension['rowIndex1'], 'col' => $this->dimension['colIndex1']],
            ['row' => $this->dimension['rowIndex2'], 'col' => $this->dimension['colIndex2']]
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
     * @param string $cellAddr
     * @param array $index
     * @param array $offset
     *
     * @return bool
     */
    protected function _validateAddressRange(&$cellAddr, &$index = null, &$offset = null)
    {
        if ($cellAddr) {
            $offset = null;
            if (is_string($cellAddr)) {
                $cellAddr = strtoupper($cellAddr);
                if (strpos($cellAddr, 'R') === 0 && strpos($cellAddr, 'C')) {
                    $offset = Excel::rangeRelOffsets($cellAddr);
                }
            }
            if ($offset && isset($offset[0], $offset[1], $offset[2], $offset[3])) {
                $dim = Excel::rangeDimensionRelative($this->dimension['cell1'], $offset);
                if (strpos($cellAddr, ':')) {
                    $cellAddr = $dim['cell1'] . ':' . $dim['cell2'];
                } else {
                    $cellAddr = $dim['cell1'];
                }
                return true;
            }
            $idxAddress = Excel::rangeIndexes($cellAddr);
            if ($idxAddress['row'] < $this->coord[0]['row'] || $idxAddress['row'] > $this->coord[1]['row']
                || $idxAddress['col'] < $this->coord[0]['col'] || $idxAddress['col'] > $this->coord[1]['col']) {
                throw new Exception('Address "' . $cellAddr . '" is outside area "' . $this->range . '"');
            }
            $index = $idxAddress;
            return true;
        }
        return false;
    }

    /**
     * setValue('A2', $value)
     * setValue([3, 1], $value) - equals to 'A1'
     * setValue('A2:C2', $value) - merge cells and write value
     *
     * @param $cellAddr
     * @param $value
     * @param $style
     *
     * @return $this
     */
    public function setValue($cellAddr, $value, $style = null)
    {
        if (is_string($cellAddr) && $this->_validateAddressRange($cellAddr)) {
            $this->sheet->setValue($cellAddr, $value, $style);
        }

        return $this;
    }

    /**
     * @param $cellAddr
     * @param $value
     * @param null $style
     *
     * @return $this
     */
    public function setFormula($cellAddr, $value, $style = null)
    {
        if (is_string($cellAddr) && $this->_validateAddressRange($cellAddr)) {
            $this->sheet->setFormula($cellAddr, $value, $style);
        }

        return $this;
    }

    /**
     * Set style (old styles wil be replaced)
     *
     * @param $cellAddr
     * @param $style
     *
     * @return $this
     */
    public function setStyle($cellAddr, $style)
    {
        if (is_string($cellAddr) && $this->_validateAddressRange($cellAddr)) {
            $this->sheet->setStyle($cellAddr, $style);
        }
        return $this;
    }

    /**
     * New style wil be added to old style (if they exists)
     *
     * @param $cellAddr
     * @param $style
     *
     * @return $this
     */
    public function addStyle($cellAddr, $style)
    {
        if (is_string($cellAddr) && $this->_validateAddressRange($cellAddr)) {
            $this->sheet->addStyle($cellAddr, $style);
        }
        return $this;
    }

    /**
     * Set style (old styles wil be replaced)
     *
     * @param $cellAddr
     * @param $format
     *
     * @return $this
     */
    public function setFormat($cellAddr, $format)
    {
        if (is_string($cellAddr) && $this->_validateAddressRange($cellAddr)) {
            $this->sheet->applayStyle($cellAddr, ['format' => $format], true);
        }
        return $this;
    }

    /**
     * @param $cellAddr
     * @param $color
     *
     * @return $this
     */
    public function setColor($cellAddr, $color)
    {
        if (is_string($cellAddr) && $this->_validateAddressRange($cellAddr)) {
            $this->sheet->applayStyle($cellAddr, ['color' => $color], true);
        }
        return $this;
    }

    /**
     * @param $cellAddr
     * @param $color
     *
     * @return $this
     */
    public function setBakgroundColor($cellAddr, $color)
    {
        if (is_string($cellAddr) && $this->_validateAddressRange($cellAddr)) {
            $this->sheet->applayStyle($cellAddr, ['fill' => $color], true);
        }
        return $this;
    }

    /**
     * @param $cellAddr
     * @param $color
     *
     * @return $this
     */
    public function setFgColor($cellAddr, $color)
    {
        return $this->setColor($cellAddr, $color);
    }

    /**
     * @param $cellAddr
     * @param $color
     *
     * @return $this
     */
    public function setBgColor($cellAddr, $color)
    {
        return $this->setBakgroundColor($cellAddr, $color);
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