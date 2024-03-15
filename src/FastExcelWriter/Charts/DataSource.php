<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Exceptions\Exception;
use avadim\FastExcelWriter\Sheet;

class DataSource
{
    const DATA_TYPE_STRING = 'string';
    const DATA_TYPE_NUMBER = 'number';

    private static array $availableDataTypes = [
        self::DATA_TYPE_STRING,
        self::DATA_TYPE_NUMBER,
    ];

    /**
     * Series Data Type
     *
     * @var string
     */
    protected string $dataType;

    /**
     * Series Data Source
     *
     * @var string
     */
    protected string $dataSource;

    /**
     * Format Code
     *
     * @var string|null
     */
    protected ?string $formatCode;

    /**
     * Point Count (The number of datapoints in the dataseries)
     *
     * @var int
     */
    protected int $pointCount = 0;


    /**
     * Create a new DataSource object
     */
    public function __construct($dataType, $dataSource = null, $formatCode = null)
    {
        $this->setDataType($dataType);
        $this->setDataSource($dataSource);
        $this->formatCode = $formatCode;
    }

    /**
     * Set Series Data Type
     *
     * @param string|null $dataType Datatype of this data series
     *                              Typical values are:
     *                              DataSeriesValues::DATA_TYPE_STRING
     *                                        Normally used for axis point values
     *                              DataSeriesValues::DATA_TYPE_NUMBER
     *                                        Normally used for chart data values
     */
    public function setDataType(string $dataType)
    {
        $dataType = strtolower($dataType);
        if (!in_array($dataType, self::$availableDataTypes)) {
            throw new Exception('Invalid datatype for chart data series values');
        }
        $this->dataType = $dataType;
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
     * Set Series Data Source (formula)
     *
     * @param string|null $dataSource
     */
    public function setDataSource(?string $dataSource = null)
    {
        $dimension = Excel::rangeDimension(str_replace('$', '', $dataSource));

        $this->dataSource = $dimension['absAddress'];
        $this->pointCount = $dimension['cellCount'];
    }

    /**
     * Get Series Data Source (formula)
     *
     * @return string
     */
    public function getDataSource(): string
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
     * Set Series Format Code
     *
     * @param string|null $formatCode
     */
    public function setFormatCode(?string $formatCode = null)
    {
        $this->formatCode = $formatCode;
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
     * Get Series Point Count
     *
     * @return int
     */
    public function getPointCount(): int
    {
        return $this->pointCount;
    }

    /**
     * @param Sheet $sheet
     * @param bool|null $force
     */
    public function applyDataSourceSheet(Sheet $sheet, ?bool $force = false)
    {
        if ($this->dataSource) {
            if ($this->dataSource[0] === '=') {
                $this->dataSource = '=' . Excel::fullAddress($sheet->getName(), substr($this->dataSource, 1), $force);
            }
            else {
                $this->dataSource = Excel::fullAddress($sheet->getName(), $this->dataSource, $force);
            }
        }
    }

}