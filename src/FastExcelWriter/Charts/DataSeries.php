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
     * Series Plot Type
     *
     * @var string
     */
    private $plotType;

    /**
     * Plot Grouping Type
     *
     * @var string|null
     */
    private ?string $plotGrouping = null;

    /**
     * Plot Direction
     *
     * @var boolean
     */
    private $plotDirection;

    /**
     * Plot Style
     *
     * @var string
     */
    private $plotStyle;

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
    private array $plotValues = [];

    /**
     * Plot Labels
     *
     * @var DataSeriesValues[] array of DataSeriesValues
     */
    private array $plotLabels = [];

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
    public function __construct($plotType = null, $plotValues = [], $plotLabels = [], $plotCategory = [], $plotGrouping = null, $plotDirection = null, $smoothLine = false, $plotStyle = null)
    {
        $this->plotType = $plotType;
        $this->plotValues = $plotValues;
        $this->plotGrouping = $plotGrouping;
        $this->plotOrder = range(0, count($plotValues) - 1);

        $this->setPlotLabels($plotLabels);

        $this->plotCategories = $plotCategory;
        $this->smoothLine = (bool)$smoothLine;
        $this->plotStyle = $plotStyle;

        if (is_null($plotDirection)) {
            $plotDirection = self::DIRECTION_COL;
        }
        $this->plotDirection = $plotDirection;
    }

    /**
     * @param Sheet $sheet
     * @param bool|null $force
     *
     * @return $this
     */
    public function applyDataSourceSheet(Sheet $sheet, ?bool $force = false): DataSeries
    {
        foreach ($this->plotValues as $plotValues) {
            if ($plotValues) {
                $plotValues->applyDataSourceSheet($sheet, $force);
            }
        }
        foreach ($this->plotLabels as $plotLabels) {
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
     * Get Plot Type
     *
     * @return string
     */
    public function getPlotType(): ?string
    {
        return $this->plotType;
    }

    /**
     * Set Plot Type
     *
     * @param string $plotType
     *
     * @return DataSeries
     */
    public function setPlotType(string $plotType = ''): DataSeries
    {
        $this->plotType = $plotType;

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
    public function getPlotDirection()
    {
        return $this->plotDirection;
    }

    /**
     * Set Plot Direction
     *
     * @param string|null $plotDirection
     *
     * @return DataSeries
     */
    public function setPlotDirection(string $plotDirection = null): DataSeries
    {
        $this->plotDirection = $plotDirection;

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
     * @param $plotLabels
     *
     * @return $this
     */
    public function setPlotLabels($plotLabels): DataSeries
    {
        foreach ($plotLabels as $n => $labels) {
            if ($labels === null) {
                $plotLabels[$n] = new DataSeriesValues();
            }
        }
        $this->plotLabels = $plotLabels;

        return $this;
    }

    /**
     * Get Plot Labels
     *
     * @return array of DataSeriesValues
     */
    public function getPlotLabels(): array
    {
        return $this->plotLabels;
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
        $keys = array_keys($this->plotLabels);
        if (in_array($index, $keys)) {
            return $this->plotLabels[$index];
        }
        elseif (isset($keys[$index])) {
            return $this->plotLabels[$keys[$index]];
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
        if (in_array($index, $keys)) {
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
     * @return string
     */
    public function getPlotStyle()
    {
        return $this->plotStyle;
    }

    /**
     * Set Plot Style
     *
     * @param string $plotStyle
     *
     * @return DataSeries
     */
    public function setPlotStyle($plotStyle = null)
    {
        $this->plotStyle = $plotStyle;

        return $this;
    }

    /**
     * Get Plot Values
     *
     * @return array of DataSeriesValues
     */
    public function getPlotValues(): array
    {
        return $this->plotValues;
    }

    /**
     * Get Plot Values by Index
     *
     * @return DataSeriesValues|false
     */
    public function getPlotValuesByIndex($index)
    {
        $keys = array_keys($this->plotValues);
        if (in_array($index, $keys)) {
            return $this->plotValues[$index];
        }
        elseif (isset($keys[$index])) {
            return $this->plotValues[$keys[$index]];
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
        return count($this->plotValues);
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