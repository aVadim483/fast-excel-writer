<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;

class DataSeriesSet
{
    private DataSeriesValues $values;
    private ?DataSeriesValues $plotLabels;
    private ?DataSeriesValues $plotCategory;

    public function __construct($dataSource, ?DataSeriesValues $plotLabels, ?DataSeriesValues $plotCategory)
    {
        $dimension = Excel::rangeDimension(str_replace('$', '', $dataSource));
        $this->values = new DataSeriesValues('Number', $dimension['absAddress'], NULL, $dimension['cellCount']);
        $this->plotLabels = $plotLabels;
        $this->plotCategory = $plotCategory;
    }
}