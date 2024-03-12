<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Exceptions\Exception;
use avadim\FastExcelWriter\Sheet;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
class DataSeriesValues
{

    const DATASERIES_TYPE_STRING = 'String';
    const DATASERIES_TYPE_NUMBER = 'Number';

    private static array $dataTypeValues = [
        self::DATASERIES_TYPE_STRING,
        self::DATASERIES_TYPE_NUMBER,
    ];

    /**
     * Series Data Type
     *
     * @var string
     */
    private string $dataType;

    /**
     * Series Data Source
     *
     * @var string
     */
    private $dataSource;

    /**
     * Format Code
     *
     * @var string
     */
    private $formatCode;

    /**
     * Series Point Marker
     *
     * @var string
     */
    private $pointMarker;

    /**
     * Point Count (The number of datapoints in the dataseries)
     *
     * @var integer
     */
    private $pointCount = 0;

    /**
     * Data Values
     *
     * @var array of mixed
     */
    private $dataValues = [];

    /**
     * Create a new DataSeriesValues object
     */
    public function __construct($dataType = self::DATASERIES_TYPE_NUMBER, $dataSource = null, $formatCode = null, $pointCount = 0, $dataValues = [], $marker = null)
    {
        $this->setDataType($dataType);
        $this->dataSource = $dataSource;
        $this->formatCode = $formatCode;
        $this->pointCount = $pointCount;
        $this->dataValues = $dataValues;
        $this->pointMarker = $marker;
    }

    /**
     * Get Series Data Type
     *
     * @return string
     */
    public function getDataType(): string
    {
        return $this->dataType;
    }

    /**
     * Set Series Data Type
     *
     * @param string|null $dataType Datatype of this data series
     *                              Typical values are:
     *                              DataSeriesValues::DATASERIES_TYPE_STRING
     *                                        Normally used for axis point values
     *                              DataSeriesValues::DATASERIES_TYPE_NUMBER
     *                                        Normally used for chart data values
     * @return $this
     */
    public function setDataType(?string $dataType = self::DATASERIES_TYPE_NUMBER): DataSeriesValues
    {
        if (!in_array($dataType, self::$dataTypeValues)) {
            throw new Exception('Invalid datatype for chart data series values');
        }
        $this->dataType = $dataType;

        return $this;
    }

    /**
     * Get Series Data Source (formula)
     *
     * @return string|null
     */
    public function getDataSource(): ?string
    {
        return $this->dataSource;
    }

    /**
     * @return bool
     */
    public function isDataSourceFormula(): bool
    {
        return $this->dataSource && $this->dataSource[0] === '=';
    }

    /**
     * Set Series Data Source (formula)
     *
     * @param string|null $dataSource
     *
     * @return $this
     */
    public function setDataSource(?string $dataSource = null): DataSeriesValues
    {
        $this->dataSource = $dataSource;

        return $this;
    }

    /**
     * Get Point Marker
     *
     * @return string
     */
    public function getPointMarker(): ?string
    {
        return $this->pointMarker;
    }

    /**
     * Set Point Marker
     *
     * @param string|null $marker
     *
     * @return $this
     */
    public function setPointMarker(?string $marker = null): DataSeriesValues
    {
        $this->pointMarker = $marker;

        return $this;
    }

    /**
     * Get Series Format Code
     *
     * @return string
     */
    public function getFormatCode(): ?string
    {
        return $this->formatCode;
    }

    /**
     * Set Series Format Code
     *
     * @param string|null $formatCode
     *
     * @return $this
     */
    public function setFormatCode(?string $formatCode = null): DataSeriesValues
    {
        $this->formatCode = $formatCode;

        return $this;
    }

    /**
     * Get Series Point Count
     *
     * @return int
     */
    public function getPointCount(): int
    {
        return $this->pointCount;
    }

    /**
     * Identify if the Data Series is a multi-level or a simple series
     *
     * @return boolean
     */
    public function isMultiLevelSeries(): ?bool
    {
        if (count($this->dataValues) > 0) {
            return is_array($this->dataValues[0]);
        }
        return null;
    }

    /**
     * Return the level count of a multi-level Data Series
     *
     * @return int
     */
    public function multiLevelCount(): int
    {
        $levelCount = 0;
        foreach ($this->dataValues as $dataValueSet) {
            $levelCount = max($levelCount, count($dataValueSet));
        }
        return $levelCount;
    }

    /**
     * Get Series Data Values
     *
     * @return array of mixed
     */
    public function getDataValues(): array
    {
        return $this->dataValues;
    }

    /**
     * Get the first Series Data value
     *
     * @return mixed
     */
    public function getDataValue()
    {
        $count = count($this->dataValues);
        if ($count == 0) {
            return null;
        }
        elseif ($count == 1) {
            return $this->dataValues[0];
        }
        return $this->dataValues;
    }

    /**
     * Set Series Data Values
     *
     * @param array $dataValues
     *
     * @return $this
     */
    public function setDataValues(array $dataValues = []): DataSeriesValues
    {
        $this->dataValues = $dataValues;
        $this->pointCount = count($dataValues);

        return $this;
    }

    /**
     * @param Sheet $sheet
     * @param bool|null $force
     *
     * @return $this
     */
    public function applyDataSourceSheet(Sheet $sheet, ?bool $force = false): DataSeriesValues
    {
        if ($this->dataSource) {
            if ($this->dataSource[0] === '=') {
                $this->dataSource = '=' . Excel::fullAddress($sheet->getName(), substr($this->dataSource, 1), $force);
            }
            else {
                $this->dataSource = Excel::fullAddress($sheet->getName(), $this->dataSource, $force);
            }
        }

        return $this;
    }
}