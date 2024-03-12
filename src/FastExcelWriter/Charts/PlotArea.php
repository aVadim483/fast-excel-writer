<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;

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


    /**
     * Create a new PlotArea
     */
    public function __construct(Layout $layout = null, $dataSeries = null)
    {
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
        }

        return $this;
    }

    /**
     * @param $dataSource
     *
     * @return $this
     */
    public function addDataSeriesSet($dataSource): PlotArea
    {
        if ($this->getPlotDataSeriesCount() === 0) {
            $this->plotDataSeries = [new DataSeries($this->defaultChartType ?: '')];
        }
        $plotDataSeries = $this->getPlotDataSeriesByIndex(0);
        $plotDataSeries->addDataSeriesValues($dataSource);
        if ($this->defaultChartType) {
            $plotDataSeries->setChartType($this->defaultChartType);
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
     * @return integer
     */
    public function getPlotSeriesCount(): int
    {
        $seriesCount = 0;
        foreach ($this->plotDataSeries as $plot) {
            if ($plot) {
                $seriesCount += $plot->getPlotSeriesCount();
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

}