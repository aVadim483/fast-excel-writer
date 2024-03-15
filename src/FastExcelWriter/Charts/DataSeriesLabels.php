<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Sheet;

class DataSeriesLabels extends DataSource
{
    /**
     * Create a new DataSeriesValues object
     */
    public function __construct($dataSource = null)
    {
        parent::__construct(self::DATA_TYPE_STRING, $dataSource);
    }

    /**
     * Set Series Data Source (formula)
     *
     * @param string|null $dataSource
     */
    public function setDataSource(?string $dataSource = null)
    {
        if ($dataSource) {
            if ($dataSource[0] === '=') {
                $dataSource = substr($dataSource, 1);
            }
            $dimension = Excel::rangeDimension(str_replace('$', '', $dataSource));
            if (isset($dimension['absAddress'])) {
                $this->dataSource = '=' . $dimension['absAddress'];
                $this->pointCount = $dimension['cellCount'];

                return;
            }
        }
        $this->dataSource = $dataSource;
        $this->pointCount = 0;
    }

    /**
     * @return bool
     */
    public function isMultiLevelSeries(): bool
    {
        return false;
    }

    /**
     * @param Sheet $sheet
     * @param bool|null $force
     *
     * @return $this
     */
    public function applyDataSourceSheet(Sheet $sheet, ?bool $force = false): DataSeriesLabels
    {
        if ($this->dataSource && $this->dataSource[0] === '=') {
            $this->dataSource = '=' . Excel::fullAddress($sheet->getName(), substr($this->dataSource, 1), $force);
        }

        return $this;
    }

}