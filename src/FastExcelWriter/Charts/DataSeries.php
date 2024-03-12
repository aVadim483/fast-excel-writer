<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Sheet;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
class DataSeries
{
    const TYPE_BARCHART = 'barChart';
    const TYPE_BARCHART_3D = 'bar3DChart';
    const TYPE_LINECHART = 'lineChart';
    const TYPE_LINECHART_3D = 'line3DChart';
    const TYPE_AREACHART = 'areaChart';
    const TYPE_AREACHART_3D = 'area3DChart';
    const TYPE_PIECHART = 'pieChart';
    const TYPE_PIECHART_3D = 'pie3DChart';
    const TYPE_DOUGHTNUTCHART = 'doughnutChart';
    const TYPE_DONUTCHART = self::TYPE_DOUGHTNUTCHART;    //    Synonym
    const TYPE_SCATTERCHART = 'scatterChart';
    const TYPE_SURFACECHART = 'surfaceChart';
    const TYPE_SURFACECHART_3D = 'surface3DChart';
    const TYPE_RADARCHART = 'radarChart';
    const TYPE_BUBBLECHART = 'bubbleChart';
    const TYPE_STOCKCHART = 'stockChart';
    const TYPE_CANDLECHART = self::TYPE_STOCKCHART;       //    Synonym

    const GROUPING_CLUSTERED = 'clustered';
    const GROUPING_STACKED = 'stacked';
    const GROUPING_PERCENT_STACKED = 'percentStacked';
    const GROUPING_STANDARD = 'standard';

    const DIRECTION_BAR = 'bar';
    const DIRECTION_HORIZONTAL = self::DIRECTION_BAR;
    const DIRECTION_COL = 'col';
    const DIRECTION_COLUMN = self::DIRECTION_COL;
    const DIRECTION_VERTICAL = self::DIRECTION_COL;

    const STYLE_LINEMARKER = 'lineMarker';
    const STYLE_SMOOTHMARKER = 'smoothMarker';
    const STYLE_MARKER = 'marker';
    const STYLE_FILLED = 'filled';


    /**
     * Series Plot Chart Type
     *
     * @var string
     */
    private string $plotChartType;

    /**
     * Plot Chart Direction
     *
     * @var string|null
     */
    private ?string $plotChartDirection = null;

    /**
     * Plot Grouping Type
     *
     * @var string|null
     */
    private ?string $plotGrouping = null;

    /**
     * Plot Style
     *
     * @var string|null
     */
    private ?string $plotStyle = null;

    /**
     * Order of plots in Series
     *
     * @var array of integer
     */
    private array $plotOrder = [];

    /**
     * Plot Values
     *
     * @var DataSeriesValues[] array of DataSeriesValues
     */
    private array $dataSeriesValues = [];

    /**
     * Plot Labels
     *
     * @var DataSeriesValues[] array of DataSeriesValues
     */
    private array $dataSeriesLabels = [];

    /**
     * Plot Category
     *
     * @var DataSeriesValues[] array of DataSeriesValues
     */
    private array $plotCategories = [];

    /**
     * Smooth Line
     *
     * @var bool
     */
    private bool $smoothLine = false;


    /**
     * Create a new DataSeries
     */
    public function __construct(string $chartType, $dataSource = null, $dataLabels = [], $plotCategories = [], $plotGrouping = null, $plotDirection = null, $smoothLine = false, $plotStyle = null)
    {
        $this->plotChartType = $chartType;
        if ($dataSource) {
            $this->setDataSeriesValues($dataSource);
        }
        if ($dataLabels) {
            $this->setDataSeriesLabels($dataLabels);
        }

        $this->plotGrouping = $plotGrouping;
        $this->plotOrder = range(0, count($this->dataSeriesValues) - 1);


        $this->plotCategories = $plotCategories;
        $this->smoothLine = (bool)$smoothLine;
        $this->plotStyle = $plotStyle;

        if (!$plotDirection) {
            $plotDirection = self::DIRECTION_COL;
        }
        $this->plotChartDirection = $plotDirection;
    }

    /**
     * @param $dataSource
     *
     * @return $this
     */
    public function setDataSeriesValues($dataSource): DataSeries
    {
        $this->dataSeriesValues = $this->dataSeriesLabels = $this->plotOrder = [];
        $this->addDataSeriesValues($dataSource);

        return $this;
    }

    /**
     * @param $dataSource
     *
     * @return $this
     */
    public function addDataSeriesValues($dataSource): DataSeries
    {
        $dataSeriesValues = [];
        $dataSeriesLabels = [];
        if ($dataSource instanceof DataSeriesValues) {
            $dataSeriesValues = [$dataSource];
        }
        elseif (is_string($dataSource)) {
            $dimension = Excel::rangeDimension(str_replace('$', '', $dataSource));
            $dataSeriesValues = [new DataSeriesValues('Number', $dimension['absAddress'], NULL, $dimension['cellCount'])];
        }
        elseif (is_array($dataSource)) {
            foreach ($dataSource as $name => $data) {
                $labelSource = (!is_int($name)) ? $name : null;
                if ($data instanceof DataSeriesValues) {
                    $dataSeriesValues[$name] = $data;
                }
                elseif (is_string($data)) {
                    $dimension = Excel::rangeDimension(str_replace('$', '', $data));
                    $dataSeriesValues[$name] = new DataSeriesValues('Number', $dimension['absAddress'], NULL, $dimension['cellCount']);
                }
                if ($labelSource) {
                    if ($labelSource[0] === '=') {
                        $labelSource = substr($labelSource, 1);
                    }
                    $dimension = Excel::rangeDimension(str_replace('$', '', $labelSource));
                    $dataSeriesLabels[$name] = new DataSeriesValues('String', '=' . $dimension['absAddress'], NULL, 1);
                }
                else {
                    $dataSeriesLabels[$name] = null;
                }
            }
        }
        if ($dataSeriesValues) {
            foreach ($dataSeriesValues as $seriesKey => $seriesValues) {
                $this->dataSeriesValues[$seriesKey] = $seriesValues;
            }
        }
        if ($dataSeriesLabels) {
            foreach ($dataSeriesLabels as $seriesKey => $seriesLabels) {
                $this->dataSeriesLabels[$seriesKey] = $seriesLabels;
            }
        }

        return $this;
    }

    /**
     * @param Sheet $sheet
     * @param bool|null $force
     *
     * @return $this
     */
    public function applyDataSourceSheet(Sheet $sheet, ?bool $force = false): DataSeries
    {
        foreach ($this->dataSeriesValues as $plotValues) {
            if ($plotValues) {
                $plotValues->applyDataSourceSheet($sheet, $force);
            }
        }
        foreach ($this->dataSeriesLabels as $plotLabels) {
            if ($plotLabels) {
                $plotLabels->applyDataSourceSheet($sheet, $force);
            }
        }
        foreach ($this->plotCategories as $plotCategories) {
            if ($plotCategories) {
                $plotCategories->applyDataSourceSheet($sheet, $force);
            }
        }

        return $this;
    }

    /**
     * @param string $chartType
     *
     * @return $this
     */
    public function setChartType(string $chartType): DataSeries
    {
        $plotChartType = $chartType;
        if (substr($chartType, -8) === '_stacked') {
            $plotChartType = str_replace('_stacked', '', $plotChartType);
        }
        if ($plotChartType === Chart::TYPE_COLUMN) {
            $plotChartType = DataSeries::TYPE_BARCHART;
        }
        elseif (substr($plotChartType, -5) !== 'Chart') {
            $plotChartType .= 'Chart';
        }

        $plotChartDirection = $plotChartGrouping = null;
        if ($chartType === Chart::TYPE_COLUMN || $chartType === Chart::TYPE_COLUMN_STACKED) {
            $plotChartDirection = DataSeries::DIRECTION_COL;
        }
        elseif (in_array($chartType, [Chart::TYPE_BAR, Chart::TYPE_BAR_STACKED])) {
            $plotChartDirection = DataSeries::DIRECTION_BAR;
        }
        if (substr($chartType, -8) === '_stacked') {
            $plotChartGrouping = DataSeries::GROUPING_STACKED;
        }
        elseif (in_array($chartType, [Chart::TYPE_BAR])) {
            $plotChartGrouping = DataSeries::GROUPING_CLUSTERED;
        }
        elseif (in_array($plotChartType, [DataSeries::TYPE_BARCHART, DataSeries::TYPE_BARCHART_3D, DataSeries::TYPE_LINECHART, DataSeries::TYPE_LINECHART_3D])) {
            $plotChartGrouping = DataSeries::GROUPING_STANDARD;
        }

        $this->setPlotChartType($plotChartType);
        if ($plotChartDirection) {
            $this->setPlotChartDirection($plotChartDirection);
        }
        if ($this->getPlotGrouping() === null && $plotChartGrouping) {
            $this->setPlotGrouping($plotChartGrouping);
        }

        return $this;
    }

    /**
     * Get Plot Type
     *
     * @return string
     */
    public function getPlotChartType(): ?string
    {
        return $this->plotChartType;
    }

    /**
     * Set Plot Type
     *
     * @param string $plotChartType
     *
     * @return DataSeries
     */
    public function setPlotChartType(string $plotChartType = ''): DataSeries
    {
        $this->plotChartType = $plotChartType;

        return $this;
    }

    /**
     * Set Plot Grouping Type
     *
     * @param string|null $groupingType
     *
     * @return DataSeries
     */
    public function setPlotGrouping(string $groupingType = null): DataSeries
    {
        $this->plotGrouping = $groupingType;

        return $this;
    }

    /**
     * Get Plot Grouping Type
     *
     * @return string
     */
    public function getPlotGrouping()
    {
        return $this->plotGrouping;
    }

    /**
     * Get Plot Direction
     *
     * @return string
     */
    public function getPlotChartDirection()
    {
        return $this->plotChartDirection;
    }

    /**
     * Set Plot Direction
     *
     * @param string|null $plotChartDirection
     *
     * @return DataSeries
     */
    public function setPlotChartDirection(string $plotChartDirection = null): DataSeries
    {
        $this->plotChartDirection = $plotChartDirection;

        return $this;
    }

    /**
     * @param array $plotOrder
     *
     * @return $this
     */
    public function setPlotOrder(array $plotOrder): DataSeries
    {
        $this->plotOrder = $plotOrder;

        return $this;
    }

    /**
     * Get Plot Order
     *
     * @return array
     */
    public function getPlotOrder(): array
    {
        return $this->plotOrder;
    }

    /**
     * @param array $dataSeriesLabels
     *
     * @return $this
     */
    public function setDataSeriesLabels(array $dataSeriesLabels): DataSeries
    {
        foreach ($dataSeriesLabels as $n => $label) {
            if ($label === null) {
                $dataSeriesLabels[$n] = new DataSeriesValues();
            }
            else {
                // PHPExcel compatible
                $source = $label->getDataSource();
                if ($source && $source[0] !== '=') {
                    $label->setDataSource('=' . $source);
                }
            }
        }
        $this->dataSeriesLabels = $dataSeriesLabels;

        return $this;
    }

    /**
     * Get Plot Labels
     *
     * @return array of DataSeriesValues
     */
    public function getDataSeriesLabels(): array
    {
        return $this->dataSeriesLabels;
    }

    /**
     * Get Plot Label by Index
     *
     * @param $index
     *
     * @return DataSeriesValues|null
     */
    public function getPlotLabelByIndex($index): ?DataSeriesValues
    {
        $keys = array_keys($this->dataSeriesLabels);
        if (in_array($index, $keys, true)) {
            return $this->dataSeriesLabels[$index];
        }
        elseif (isset($keys[$index])) {
            return $this->dataSeriesLabels[$keys[$index]];
        }
        return null;
    }

    /**
     * @param $categories
     *
     * @return $this
     */
    public function setPlotCategories($categories): DataSeries
    {
        $plotCategories = [];
        if ($categories instanceof DataSeriesValues) {
            $plotCategories = [$categories];
        }
        elseif (is_array($categories)) {
            foreach ($categories as $category) {
                if ($category instanceof DataSeriesValues) {
                    $plotCategories[] = $category;
                }
                elseif (is_string($category)) {
                    $dimension = Excel::rangeDimension(str_replace('$', '', $category));
                    $plotCategories[] = new DataSeriesValues('String', $dimension['absAddress'], NULL, $dimension['cellCount']);
                }
            }
        }
        elseif (is_string($categories)) {
            $dimension = Excel::rangeDimension(str_replace('$', '', $categories));
            $plotCategories = [new DataSeriesValues('String', $dimension['absAddress'], NULL, $dimension['cellCount'])];
        }
        $this->plotCategories = $plotCategories;

        return $this;
    }

    /**
     * Get Plot Categories
     *
     * @return array of DataSeriesValues
     */
    public function getPlotCategories(): array
    {
        return $this->plotCategories;
    }

    /**
     * Get Plot Category by Index
     *
     * @param $index
     *
     * @return DataSeriesValues|null
     */
    public function getPlotCategoryByIndex($index): ?DataSeriesValues
    {
        $keys = array_keys($this->plotCategories);
        if (in_array($index, $keys, true)) {
            return $this->plotCategories[$index];
        }
        elseif (isset($keys[$index])) {
            return $this->plotCategories[$keys[$index]];
        }

        return null;
    }

    /**
     * Get Plot Style
     *
     * @return string|null
     */
    public function getPlotStyle(): ?string
    {
        return $this->plotStyle;
    }

    /**
     * Set Plot Style
     *
     * @param string|null $plotStyle
     *
     * @return DataSeries
     */
    public function setPlotStyle(?string $plotStyle = null): DataSeries
    {
        $this->plotStyle = $plotStyle;

        return $this;
    }

    /**
     * Get Plot Values
     *
     * @return array of DataSeriesValues
     */
    public function getDataSeriesValues(): array
    {
        return $this->dataSeriesValues;
    }

    /**
     * Get Plot Values by Index
     *
     * @return DataSeriesValues|false
     */
    public function getPlotValuesByIndex($index)
    {
        $keys = array_keys($this->dataSeriesValues);
        if (in_array($index, $keys, true)) {
            return $this->dataSeriesValues[$index];
        }
        elseif (isset($keys[$index])) {
            return $this->dataSeriesValues[$keys[$index]];
        }
        return false;
    }

    /**
     * Get Number of Plot Series
     *
     * @return int
     */
    public function getPlotSeriesCount(): int
    {
        return count($this->dataSeriesValues);
    }

    /**
     * Get Smooth Line
     *
     * @return bool
     */
    public function getSmoothLine(): bool
    {
        return $this->smoothLine;
    }

    /**
     * Set Smooth Line
     *
     * @param bool $smoothLine
     *
     * @return DataSeries
     */
    public function setSmoothLine(bool $smoothLine = true): DataSeries
    {
        $this->smoothLine = $smoothLine;

        return $this;
    }

}