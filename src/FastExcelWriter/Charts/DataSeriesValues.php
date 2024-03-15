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
class DataSeriesValues extends DataSource
{
    /**
     * Series Point Marker
     *
     * @var string|null
     */
    private ?string $pointMarker = null;

    /**
     * Data Values
     *
     * @var array of mixed
     */
    private array $dataValues = [];

    /**
     * @var DataSeriesLabels|null
     */
    private ?DataSeriesLabels $labels = null;

    /**
     * @var array
     */
    private array $options = [];


    /**
     * Create a new DataSeriesValues object
     */
    public function __construct($dataSource = null, $dataLabels = null, $options = [])
    {
        parent::__construct(self::DATA_TYPE_NUMBER, $dataSource);
        $this->setLabels($dataLabels);
        $this->setOptions($options);
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
        if ($count === 0) {
            return null;
        }
        elseif ($count === 1) {
            return $this->dataValues[0];
        }
        return $this->dataValues;
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
     * @param DataSeriesLabels|string $dataLabels
     *
     * @return void
     */
    public function setLabels($dataLabels)
    {
        if ($dataLabels instanceof DataSeriesLabels) {
            $this->labels = $dataLabels;
        }
        elseif (is_string($dataLabels)) {
            $this->labels = new DataSeriesLabels($dataLabels);
        }
    }

    /**
     * @return DataSeriesLabels|null
     */
    public function getLabels(): ?DataSeriesLabels
    {
        return $this->labels;
    }

    /**
     * @param array|null $options
     *
     * @return void
     */
    public function setOptions(?array $options = [])
    {
        foreach ($options as $key => $val) {
            switch ($key) {
                case 'color':
                    $this->setColor($val);
                    break;
                case 'marker':
                    $this->setPointMarker($val);
                    break;
                case 'format':
                    $this->setFormatCode($val);
                    break;
                default:
                    //
            }
        }
        $this->options = $options;
    }

    /**
     * @return array
     */
    public function getOptions(): array
    {
        return $this->options;
    }

    public function setColor($color)
    {
        if (preg_match('/^#?([0-9a-f]+)$/i', $color, $m)) {
            $this->options['color'] = $m[1];
        }
    }

    public function getColor()
    {
        return $this->options['color'] ?? null;
    }

}