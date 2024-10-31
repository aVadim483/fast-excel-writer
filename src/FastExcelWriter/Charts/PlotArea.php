<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Exceptions\Exception;
use avadim\FastExcelWriter\Exceptions\ExceptionChart;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
class PlotArea
{
    /**
     * PlotArea Layout
     *
     * @var Layout|null
     */
    private ?Layout $layout = null;

    /**
     * Plot Series
     *
     * @var DataSeries[] array of DataSeries
     */
    private array $plotDataSeries = [];

    private ?string $defaultChartType = null;

    private array $defaultColors = [];


    /**
     * Create a new PlotArea
     */
    public function __construct($dataSeries = null, ?Layout $layout = null)
    {
        $this->defaultColors = ['5b9bd5', 'ed7d31', 'a5a5a5', 'ffc000', '4472c4', '70ad47'];
        $this->layout = $layout;
        if ($dataSeries) {
            if ($dataSeries instanceof DataSeries) {
                $this->plotDataSeries = [$dataSeries];
            }
            elseif (is_array($dataSeries) && current($dataSeries) instanceof DataSeries) {
                $this->plotDataSeries = $dataSeries;
            }
            else {
                $this->addDataSeriesSet($dataSeries);
            }
        }
    }

    /**
     * @param string $chartType
     *
     * @return $this
     */
    public function setChartType(string $chartType): PlotArea
    {
        $this->defaultChartType = $chartType;
        foreach ($this->getPlotDataSeries() as $plotSeries) {
            $plotSeries->setChartType($chartType);
            if (in_array($plotSeries->getChartType(), [Chart::TYPE_PIE, Chart::TYPE_PIE_3D, Chart::TYPE_DONUT])) {
                $dataSeriesValues = $plotSeries->getDataSeriesValues();
                if (!$dataSeriesValues[0]->getSegmentColors()) {
                    $dataSeriesValues[0]->setSegmentColors($this->defaultColors);
                }
            }
        }

        return $this;
    }

    /**
     * @param array $colors
     *
     * @return $this
     */
    public function setDefaultColors(array $colors): PlotArea
    {
        foreach (array_values($colors) as $n => $color) {
            $this->defaultColors[$n] = $color;
        }

        return $this;
    }

    /**
     * @return string|null
     */
    protected function selectDefaultColor(): ?string
    {
        $index = $this->getDataSeriesCount();

        return $this->defaultColors[$index] ?? null;
    }

    /**
     * @param $dataSource
     * @param string|null $dataLabel
     * @param array|null $options
     *
     * @return $this
     */
    public function addDataSeriesValues($dataSource, ?string $dataLabel = null, ?array $options = []): PlotArea
    {
        $chartType = $options['chart_type'] ?? $this->defaultChartType;
        if ($this->getPlotDataSeriesCount() === 0) {
            //$this->plotDataSeries = [new DataSeries($chartType)];
        }
        $axisNum = $options['axis_num'] ?? 1;

        $key = $chartType . '-' . $axisNum;
        if (isset($this->plotDataSeries[$key])) {
            $dataSeries = $this->plotDataSeries[$key];
        }
        else {
            $dataSeries = new DataSeries($chartType);
            $dataSeries->setAxisNum($axisNum);
            $this->plotDataSeries[$key] = $dataSeries;
        }

        if (!isset($options['color']) && ($color = $this->selectDefaultColor())) {
            $options['color'] = $color;
        }
        $dataSeries->addDataSeriesValues($dataSource, $dataLabel, $options);
        if ($this->defaultChartType) {
            $dataSeries->setChartType($chartType);
        }

        return $this;
    }

    /**
     * @param array $dataSources
     *
     * @return $this
     */
    public function addDataSeriesSet(array $dataSources): PlotArea
    {
        foreach ($dataSources as $name => $dataSource) {
            $this->addDataSeriesValues($dataSource, $name);
        }

        return $this;
    }

    /**
     * @param $layout
     *
     * @return $this
     */
    public function setLayout($layout): PlotArea
    {
        $this->layout = $layout;

        return $this;
    }

    /**
     * Get Layout
     *
     * @return Layout
     */
    public function getLayout(): ?Layout
    {
        return $this->layout;
    }

    /**
     * Get Number of Plot Groups
     *
     * @return int of DataSeries
     */
    public function getPlotDataSeriesCount(): int
    {
        return count($this->plotDataSeries);
    }

    /**
     * Get Number of Plot Series
     *
     * @return int
     */
    public function getDataSeriesCount(): int
    {
        $seriesCount = 0;
        foreach ($this->plotDataSeries as $plot) {
            if ($plot) {
                $seriesCount += $plot->getDataSeriesCount();
            }
        }
        return $seriesCount;
    }

    /**
     * Get array of DataSeries
     *
     * @return DataSeries[] array of DataSeries
     */
    public function getPlotDataSeries(): array
    {
        return $this->plotDataSeries;
    }

    /**
     * Get DataSeries by Index
     *
     * @param $index
     *
     * @return DataSeries
     */
    public function getPlotDataSeriesByIndex($index): ?DataSeries
    {
        $keys = array_keys($this->plotDataSeries);
        if (in_array($index, $keys, true)) {
            return $this->plotDataSeries[$index];
        }
        elseif (isset($keys[$index])) {
            return $this->plotDataSeries[$keys[$index]];
        }
        return null;
    }

    /**
     * @return array
     */
    public function getChartTypes(): array
    {
        $chartTypes = [];
        foreach ($this->plotDataSeries as $dataSeries) {
            if ($type = $dataSeries->getPlotChartType()) {
                $chartTypes[] = $type;
            }
        }
        if (!$chartTypes) {
            throw new ExceptionChart('Chart is not yet implemented');
        }

        return $chartTypes;
    }
}