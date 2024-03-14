<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Sheet;

class DataSeriesLabels extends DataSeriesValues
{
    /**
     * @param Sheet $sheet
     * @param bool|null $force
     *
     * @return $this
     */
    public function applyDataSourceSheet(Sheet $sheet, ?bool $force = false): DataSeriesValues
    {
        if ($this->dataSource && $this->dataSource[0] === '=') {
            $this->dataSource = '=' . Excel::fullAddress($sheet->getName(), substr($this->dataSource, 1), $force);
        }

        return $this;
    }

}