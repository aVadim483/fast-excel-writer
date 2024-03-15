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
    public function __construct($dataSeries = null, Layout $layout = null)
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
     * @return string|null
     */
    protected function getDefaultColor(): ?string
    {
        static $colors = ['5b9bd5', 'ed7d31', 'a5a5a5', 'ffc000', '4472c4', '70ad47'];

        $index = $this->getDataSeriesCount();

        return $colors[$index] ?? null;
    }

    /**
     * @param $dataSource
     * @param string|null $dataLabel
     * @param array|null $options
     *
     * @return $this
     */
    public function addDataSeriesSource($dataSource, ?string $dataLabel = null, ?array $options = []): PlotArea
    {
        if ($this->getPlotDataSeriesCount() === 0) {
            $this->plotDataSeries = [new DataSeries($this->defaultChartType ?: '')];
        }
        $plotDataSeries = $this->getPlotDataSeriesByIndex(0);
        if (!isset($options['color']) && ($color = $this->getDefaultColor())) {
            $options['color'] = $color;
        }
        $plotDataSeries->addDataSeriesSource($dataSource, $dataLabel, $options);
        if ($this->defaultChartType) {
            $plotDataSeries->setChartType($this->defaultChartType);
        }

        return $this;
    }

    /**
     * @param array $dataSource
     *
     * @return $this
     */
    public function addDataSeriesSet(array $dataSource): PlotArea
    {
        foreach ($dataSource as $name => $values) {
            $this->addDataSeriesSource($values, $name);
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

}