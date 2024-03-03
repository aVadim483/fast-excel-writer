<?php

namespace avadim\FastExcelWriter\Charts;

class PlotArea
{
    /**
     * PlotArea Layout
     *
     * @var Layout|null
     */
    private $layout = null;

    /**
     * Plot Series
     *
     * @var array of DataSeries
     */
    private $plotSeries = array();

    /**
     * Create a new PlotArea
     */
    public function __construct(Layout $layout = null, $plotSeries = array())
    {
        $this->layout = $layout;
        $this->plotSeries = $plotSeries;
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
    public function getPlotGroupCount()
    {
        return count($this->plotSeries);
    }

    /**
     * Get Number of Plot Series
     *
     * @return integer
     */
    public function getPlotSeriesCount()
    {
        $seriesCount = 0;
        foreach ($this->plotSeries as $plot) {
            $seriesCount += $plot->getPlotSeriesCount();
        }
        return $seriesCount;
    }

    /**
     * Get Plot Series
     *
     * @return array of DataSeries
     */
    public function getPlotGroup()
    {
        return $this->plotSeries;
    }

    /**
     * Get Plot Series by Index
     *
     * @return DataSeries
     */
    public function getPlotGroupByIndex($index)
    {
        return $this->plotSeries[$index];
    }

    /**
     * Set Plot Series
     *
     * @param [DataSeries]
     * @return PlotArea
     */
    public function setPlotSeries($plotSeries = array())
    {
        $this->plotSeries = $plotSeries;

        return $this;
    }


}