<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Exceptions\Exception;
use avadim\FastExcelWriter\Exceptions\ExceptionChart;
use avadim\FastExcelWriter\Sheet;
use avadim\FastExcelWriter\Writer\Writer;
use avadim\FastExcelWriter\Writer\FileWriter;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
class Chart
{
    const TYPE_BAR              = 'bar';
    const TYPE_BAR_STACKED      = 'bar_stacked';
    const TYPE_COLUMN           = 'column';
    const TYPE_COLUMN_STACKED   = 'column_stacked';
    const TYPE_LINE             = 'line';
    const TYPE_LINE_STACKED     = 'line_stacked';
    const TYPE_LINE_3D          = 'line3D';
    const TYPE_LINE_3D_STACKED  = 'line3D_stacked';
    const TYPE_AREA             = 'area';
    const TYPE_AREA_STACKED     = 'area_stacked';
    const TYPE_AREA_3D          = 'area3D';
    const TYPE_AREA_3D_STACKED  = 'area3D_stacked';
    const TYPE_PIE              = 'pie';
    const TYPE_PIE_3D           = 'pie3D';
    const TYPE_DONUT            = 'doughnut';
    const TYPE_DOUGHNUT         = self::TYPE_DONUT;
    const TYPE_COMBO            = 'combo';

    public static array $charTypes = [
        self::TYPE_BAR              ,
        self::TYPE_BAR_STACKED      ,
        self::TYPE_COLUMN           ,
        self::TYPE_COLUMN_STACKED   ,
        self::TYPE_LINE             ,
        self::TYPE_LINE_STACKED     ,
        self::TYPE_LINE_3D          ,
        self::TYPE_LINE_3D_STACKED  ,
        self::TYPE_AREA             ,
        self::TYPE_AREA_STACKED     ,
        self::TYPE_AREA_3D          ,
        self::TYPE_AREA_3D_STACKED  ,
        self::TYPE_PIE              ,
        self::TYPE_PIE_3D           ,
        self::TYPE_DONUT            ,
        self::TYPE_DOUGHNUT         ,
        self::TYPE_COMBO            ,
    ];

    public string $fileName;

    public string $rId;

    private string $chartType;

    /**
     * @var Sheet|null
     */
    private ?Sheet $sheet = null;

    /**
     * Chart Name
     *
     * @var string
     */
    private string $name = '';

    /**
     * Chart Title
     *
     * @var Title
     */
    private Title $title;

    /**
     * Chart Legend
     *
     * @var Legend|null
     */
    private ?Legend $legend;

    /**
     * Category Axis Title
     *
     * @var Title|null
     */
    private $categoryAxisTitle = null;

    /**
     * Y-Axis Title
     *
     * @var Title|null
     */
    private $valueAxisTitle = null;

    /**
     * Y2-Axis Title
     *
     * @var Title|null
     */
    private $valueAxisTitle2 = null;

    /**
     * Chart Plot Area
     *
     * @var PlotArea
     */
    private PlotArea $plotArea;

    /**
     * Plot Visible Only
     *
     * @var bool
     */
    private bool $plotVisibleOnly = true;

    /**
     * Display Blanks as
     *
     * @var string
     */
    private $displayBlanksAs = '0';

    /**
     * Chart Axis X as
     *
     * @var Axis|null
     */
    private ?Axis $xAxis;

    /**
     * Chart Axis Y as
     *
     * @var Axis|null
     */
    private ?Axis $yAxis;

    /**
     * Chart Axis Y2 as
     *
     * @var Axis|null
     */
    private ?Axis $yAxis2;

    /**
     * Chart Major Gridlines as
     *
     * @var GridLines|null
     */
    private ?GridLines $majorGridlines;

    /**
     * Chart Minor Gridlines as
     *
     * @var GridLines|null
     */
    private ?GridLines $minorGridlines;

    /**
     * Top-Left Cell Position
     *
     * @var string
     */
    private string $topLeftCellRef = 'A1';


    /**
     * Top-Left X-Offset
     *
     * @var int
     */
    private int $topLeftXOffset = 0;


    /**
     * Top-Left Y-Offset
     *
     * @var int
     */
    private int $topLeftYOffset = 0;


    /**
     * Bottom-Right Cell Position
     *
     * @var string
     */
    private string $bottomRightCellRef = 'A1';


    /**
     * Bottom-Right X-Offset
     *
     * @var int
     */
    private int $bottomRightXOffset = 10;


    /**
     * Bottom-Right Y-Offset
     *
     * @var int
     */
    private int $bottomRightYOffset = 10;
    
    private int $_seriesIndex;


    /**
     * Chart constructor
     *
     * @param Title|string $title
     * @param PlotArea|array $plotArea
     * @param Legend|null $legend
     * @param bool|null $plotVisibleOnly
     * @param string|null $displayBlanksAs
     * @param Title|string|null $xAxisLabel
     * @param Title|string|null $yAxisLabel
     * @param Axis|null $xAxis
     * @param Axis|null $yAxis
     * @param GridLines|null $majorGridlines
     * @param GridLines|null $minorGridlines
     */
    public function __construct($title, $plotArea, ?Legend $legend = null, ?bool $plotVisibleOnly = true, ?string $displayBlanksAs = '0',
                                $xAxisLabel = null, $yAxisLabel = null, ?Axis $xAxis = null, ?Axis $yAxis = null, ?GridLines $majorGridlines = null,
                                ?GridLines $minorGridlines = null)
    {
        $this->setTitle($title);
        $this->legend = $legend;
        if ($xAxisLabel) {
            $this->categoryAxisTitle = is_string($xAxisLabel) ? (new Title($xAxisLabel)) : $xAxisLabel;
        }
        if ($yAxisLabel) {
            $this->valueAxisTitle = is_string($yAxisLabel) ? (new Title($yAxisLabel)) : $yAxisLabel;
        }

        $this->setPlotArea($plotArea);
        $this->plotVisibleOnly = $plotVisibleOnly;
        $this->displayBlanksAs = $displayBlanksAs;
        $this->xAxis = $xAxis;
        $this->yAxis = $yAxis;
        $this->majorGridlines = $majorGridlines;
        $this->minorGridlines = $minorGridlines;
    }

    /**
     * @param string $chartType
     * @param Title|string $title
     * @param DataSeries|array $dataSource
     *
     * @return $this
     */
    public static function make(string $chartType, $title = null, $dataSource = null): Chart
    {
        if (!in_array($chartType, self::$charTypes)) {
            ExceptionChart::throwNew('Invalid chart type "' . $chartType . '"');
        }
        if ($dataSource instanceof PlotArea) {
            $plotArea = $dataSource;
        }
        else {
            if (in_array($chartType, [Chart::TYPE_PIE, Chart::TYPE_PIE_3D, Chart::TYPE_DONUT])) {
                $plotArea = new PlotArea(new DataSeries($chartType, $dataSource));
            }
            else {
                $plotArea = new PlotArea($dataSource);
            }
        }
        $plotArea->setChartType($chartType);
        $chart = new static($title, $plotArea);
        $chart->setChartType($chartType);

        return $chart;
    }


    /**
     * @param $dataSource
     * @param string|null $dataLabel
     * @param array|null $options
     *
     * @return $this
     */
    public function addDataSeriesValues($dataSource, ?string $dataLabel = null, ?array $options = []): Chart
    {
        if ($this->chartType === Chart::TYPE_COMBO) {
            ExceptionChart::throwNew('Please use method addDataSeriesType() for Chart "' . Chart::TYPE_COMBO . '"');
        }
        $this->getPlotArea()->addDataSeriesValues($dataSource, $dataLabel, $options);

        return $this;
    }

    /**
     * @param array $dataSources
     *
     * @return $this
     */
    public function addDataSeriesSet(array $dataSources): Chart
    {
        $this->getPlotArea()->addDataSeriesSet($dataSources);

        return $this;
    }

    /**
     * @param string $chartType
     * @param DataSeriesValues|string $dataSource
     * @param string|null $dataLabel
     * @param array|null $options
     *
     * @return $this
     */
    public function addDataSeriesType(string $chartType, $dataSource, ?string $dataLabel = null, ?array $options = []): Chart
    {
        if ($this->chartType === self::TYPE_COMBO && !in_array($chartType, [self::TYPE_COLUMN, self::TYPE_LINE])) {
            ExceptionChart::throwNew('Invalid chart type of DataSeries "' . $chartType . '" for Chart "' . self::TYPE_COMBO . '"');
        }
        elseif ($this->chartType !== self::TYPE_COMBO && $this->chartType !== $chartType) {
            ExceptionChart::throwNew('Invalid chart type of DataSeries "' . $chartType . '" for Chart "' . $this->chartType . '"');
        }
        $options['chart_type'] = $chartType;
        $this->getPlotArea()->addDataSeriesValues($dataSource, $dataLabel, $options);

        return $this;
    }

    /**
     * @param string $chartType
     * @param DataSeriesValues|string $dataSource
     * @param string|null $dataLabel
     * @param array|null $options
     *
     * @return $this
     */
    public function addDataSeriesType2(string $chartType, $dataSource, ?string $dataLabel = null, ?array $options = []): Chart
    {
        $options['axis_num'] = 2;

        return $this->addDataSeriesType($chartType, $dataSource, $dataLabel, $options);
    }

    /**
     * @param string $name
     *
     * @return $this
     */
    public function setName(string $name): Chart
    {
        $this->name = $name;

        return $this;
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName(): string
    {
        return $this->name;
    }

    /**
     * Set Title
     *
     * @param Title|string $title
     *
     * @return $this
     */
    public function setTitle($title): Chart
    {
        if ($title instanceof Title) {
            $this->title = $title;
        }
        else {
            $this->title = new Title($title);
        }

        return $this;
    }

    /**
     * Get Title
     *
     * @return Title
     */
    public function getTitle(): Title
    {
        return $this->title;
    }

    /**
     * @param $plotValues
     *
     * @return $this
     */
    public function setPlotArea($plotValues): Chart
    {
        if ($plotValues instanceof PlotArea) {
            $this->plotArea = $plotValues;
        }
        else {
            $this->plotArea = new PlotArea($plotValues);
        }

        return $this;
    }

    /**
     * @param $layout
     *
     * @return $this
     */
    public function setPlotLayout($layout): Chart
    {
        $this->plotArea->setLayout($layout);

        return $this;
    }

    /**
     * @param bool $val
     *
     * @return $this
     */
    public function setPlotShowValues(bool $val): Chart
    {
        $layout = $this->plotArea->getLayout();
        if (!$layout) {
            $layout = new Layout();
        }
        $layout->setShowVal($val);
        $this->plotArea->setLayout($layout);

        return $this;
    }

    /**
     * @param bool $val
     *
     * @return $this
     */
    public function setPlotShowPercent(bool $val): Chart
    {
        $layout = $this->plotArea->getLayout();
        if (!$layout) {
            $layout = new Layout();
        }
        $layout->setShowPercent($val);
        $this->plotArea->setLayout($layout);

        return $this;
    }

    /**
     * Get Legend
     *
     * @return Legend
     */
    public function getLegend(): ?Legend
    {
        return $this->legend;
    }

    /**
     * Set Legend
     *
     * @param Legend $legend
     * 
     * @return $this
     */
    public function setLegend(Legend $legend): Chart
    {
        $this->legend = $legend;

        return $this;
    }

    /**
     * @param string $position
     *
     * @return $this
     */
    public function setLegendPosition(string $position): Chart
    {
        if (!$this->legend) {
            $this->legend = new Legend($position);
        }
        else {
            $this->legend->setPosition($position);
        }

        return $this;
    }

    /**
     * @return $this
     */
    public function setLegendPositionTop(): Chart
    {
        if (!$this->legend) {
            $this->legend = new Legend(Legend::POSITION_TOP);
        }
        else {
            $this->legend->setPosition(Legend::POSITION_TOP);
        }

        return $this;
    }

    /**
     * @return $this
     */
    public function setLegendPositionRight(): Chart
    {
        if (!$this->legend) {
            $this->legend = new Legend(Legend::POSITION_RIGHT);
        }
        else {
            $this->legend->setPosition(Legend::POSITION_RIGHT);
        }

        return $this;
    }

    /**
     * @return $this
     */
    public function setLegendPositionBottom(): Chart
    {
        if (!$this->legend) {
            $this->legend = new Legend(Legend::POSITION_BOTTOM);
        }
        else {
            $this->legend->setPosition(Legend::POSITION_BOTTOM);
        }

        return $this;
    }

    /**
     * @return $this
     */
    public function setLegendPositionLeft(): Chart
    {
        if (!$this->legend) {
            $this->legend = new Legend(Legend::POSITION_LEFT);
        }
        else {
            $this->legend->setPosition(Legend::POSITION_LEFT);
        }

        return $this;
    }

    /**
     * Set Category Axis Labels
     *
     * @param $labels
     *
     * @return $this
     */
    public function setCategoryAxisLabels($labels): Chart
    {
        $dataSeries = $this->plotArea->getPlotDataSeriesByIndex(0);
        $dataSeries->setCategoryAxisLabels($labels);

        return $this;
    }

    /**
     * @param $labels
     * @param $title
     *
     * @return $this
     */
    public function setCategoryAxis($labels, $title = null): Chart
    {
        $this->setCategoryAxisLabels($labels);
        if ($title) {
            $this->setCategoryAxisTitle($title);
        }

        return $this;
    }

    /**
     * Backward compatible
     * @deprecated

     * @param $range
     *
     * @return $this
     */
    public function setDataSeriesTickLabels($range): Chart
    {

        return $this->setCategoryAxisLabels($range);
    }

    /**
     * Set Category Axis Title
     *
     * @param Title|string $title
     *
     * @return $this
     */
    public function setCategoryAxisTitle($title): Chart
    {
        $this->categoryAxisTitle = is_string($title) ? (new Title($title)) : $title;

        return $this;
    }

    /**
     * Get Category Axis Title
     *
     * @return Title
     */
    public function getCategoryAxisTitle(): ?Title
    {
        return $this->categoryAxisTitle;
    }

    /**
     * Backward compatible
     * @deprecated
     *
     * @param Title|string $title
     *
     * @return $this
     */
    public function setXAxisLabel($title): Chart
    {

        return $this->setCategoryAxisTitle($title);
    }

    /**
     * Backward compatible
     * @deprecated
     *
     * @return Title
     */
    public function getXAxisLabel(): ?Title
    {
        return $this->getCategoryAxisTitle();
    }


    /**
     * Set Value Axis Title
     *
     * @param Title|string $title
     *
     * @return $this
     */
    public function setValueAxisTitle($title): Chart
    {
        $this->valueAxisTitle = is_string($title) ? (new Title($title)) : $title;

        return $this;
    }

    /**
     * Get Value Axis Title
     *
     * @return Title
     */
    public function getValueAxisTitle(): ?Title
    {
        return $this->valueAxisTitle;
    }

    /**
     * Set Y2 Axis Title
     *
     * @param Title|string $title
     *
     * @return $this
     */
    public function setValueAxisTitle2($title): Chart
    {
        $this->valueAxisTitle2 = is_string($title) ? (new Title($title)) : $title;

        return $this;
    }

    /**
     * Get Y2 Axis Title
     *
     * @return Title
     */
    public function getValueAxisTitle2(): ?Title
    {
        return $this->valueAxisTitle2;
    }

    /**
     * Backward compatible
     * @deprecated
     *
     * @param Title|string $title
     *
     * @return $this
     */
    public function setYAxisLabel($title): Chart
    {

        return $this->setValueAxisTitle($title);
    }

    /**
     * Backward compatible
     * @deprecated
     *
     * @return Title
     */
    public function getYAxisLabel(): ?Title
    {
        return $this->getValueAxisTitle();
    }

    public function setDataSeriesNames($labels): Chart
    {
        $dataSeries = $this->plotArea->getPlotDataSeriesByIndex(0);
        $dataSeries->setDataSeriesLabels($labels);

        return $this;
    }

    /**
     * Get Plot Area
     *
     * @return PlotArea
     */
    public function getPlotArea(): PlotArea
    {
        return $this->plotArea;
    }

    /**
     * Get Plot Visible Only
     *
     * @return boolean
     */
    public function getPlotVisibleOnly(): bool
    {
        return $this->plotVisibleOnly;
    }

    /**
     * Set Plot Visible Only
     *
     * @param boolean $plotVisibleOnly
     *
     * @return $this
     */
    public function setPlotVisibleOnly(bool $plotVisibleOnly = true): Chart
    {
        $this->plotVisibleOnly = $plotVisibleOnly;

        return $this;
    }

    /**
     * Get Display Blanks as
     *
     * @return string
     */
    public function getDisplayBlanksAs(): string
    {
        return $this->displayBlanksAs;
    }

    /**
     * Set Display Blanks as
     *
     * @param string $displayBlanksAs
     */
    public function setDisplayBlanksAs(string $displayBlanksAs = '0')
    {
        $this->displayBlanksAs = $displayBlanksAs;
    }

    /**
     * Get xAxis
     *
     * @return Axis
     */
    public function getChartAxisX(): ?Axis
    {
        if ($this->xAxis !== null) {
            return $this->xAxis;
        }

        return new Axis();
    }

    /**
     * Get yAxis
     *
     * @return Axis
     */
    public function getChartAxisY(): ?Axis
    {
        if ($this->yAxis !== null) {
            return $this->yAxis;
        }

        return new Axis();
    }

    /**
     * Get yAxis
     *
     * @return Axis
     */
    public function getChartAxisY2(): ?Axis
    {
        if ($this->yAxis !== null) {
            return $this->yAxis;
        }

        return new Axis();
    }

    /**
     * Get Major Gridlines
     *
     * @return GridLines
     */
    public function getMajorGridlines(): ?GridLines
    {
        if ($this->majorGridlines !== null) {
            return $this->majorGridlines;
        }

        return new GridLines();
    }

    /**
     * Get Minor Gridlines
     *
     * @return GridLines
     */
    public function getMinorGridlines(): ?GridLines
    {
        if ($this->minorGridlines !== null) {
            return $this->minorGridlines;
        }

        return new GridLines();
    }


    /**
     * Set the Top Left position for the chart
     *
     * @param string $cell
     * @param integer|null $xOffset
     * @param integer|null $yOffset
     *
     * @return $this
     */
    public function setTopLeftPosition(string $cell, ?int $xOffset = null, ?int $yOffset = null): Chart
    {
        $this->topLeftCellRef = $cell;
        if (!is_null($xOffset)) {
            $this->setTopLeftXOffset($xOffset);
        }
        if (!is_null($yOffset)) {
            $this->setTopLeftYOffset($yOffset);
        }

        return $this;
    }

    /**
     * Get the top left position of the chart
     *
     * @return array an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getTopLeftPosition(): array
    {
        return [
            'cell'    => $this->topLeftCellRef,
            'xOffset' => $this->topLeftXOffset,
            'yOffset' => $this->topLeftYOffset
        ];
    }

    /**
     * Get the cell address where the top left of the chart is fixed
     *
     * @return string
     */
    public function getTopLeftCell(): string
    {
        return $this->topLeftCellRef;
    }

    /**
     * Set the Top Left cell position for the chart
     *
     * @param string $cell
     *
     * @return $this
     */
    public function setTopLeftCell(string $cell): Chart
    {
        $this->topLeftCellRef = $cell;

        return $this;
    }

    /**
     * Set the offset position within the Top Left cell for the chart
     *
     * @param integer|null $xOffset
     * @param integer|null $yOffset
     *
     * @return $this
     */
    public function setTopLeftOffset(?int $xOffset = null, ?int $yOffset = null): Chart
    {
        if (null !== $xOffset) {
            $this->setTopLeftXOffset($xOffset);
        }
        if (null !== $yOffset) {
            $this->setTopLeftYOffset($yOffset);
        }

        return $this;
    }

    /**
     * Get the offset position within the Top Left cell for the chart
     *
     * @return int[]
     */
    public function getTopLeftOffset(): array
    {
        return [
            'X' => $this->topLeftXOffset,
            'Y' => $this->topLeftYOffset
        ];
    }

    /**
     * @param $xOffset
     *
     * @return $this
     */
    public function setTopLeftXOffset($xOffset): Chart
    {
        $this->topLeftXOffset = $xOffset;

        return $this;
    }

    /**
     * @return int
     */
    public function getTopLeftXOffset(): int
    {
        return $this->topLeftXOffset;
    }

    /**
     * @param $yOffset
     *
     * @return $this
     */
    public function setTopLeftYOffset($yOffset): Chart
    {
        $this->topLeftYOffset = $yOffset;

        return $this;
    }

    /**
     * @return int
     */
    public function getTopLeftYOffset(): int
    {
        return $this->topLeftYOffset;
    }

    /**
     * Set the Bottom Right position of the chart
     *
     * @param string $cell
     * @param int|null $xOffset
     * @param int|null $yOffset
     *
     * @return $this
     */
    public function setPosition(string $cell, ?int $xOffset = null, ?int $yOffset = null): Chart
    {
        $this->bottomRightCellRef = $cell;
        if ($xOffset !== null) {
            $this->setBottomRightXOffset($xOffset);
        }
        if ($yOffset !== null) {
            $this->setBottomRightYOffset($yOffset);
        }

        return $this;
    }

    /**
     * Get the bottom right position of the chart
     *
     * @return array an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getBottomRightPosition(): array
    {
        return array(
            'cell'    => $this->bottomRightCellRef,
            'xOffset' => $this->bottomRightXOffset,
            'yOffset' => $this->bottomRightYOffset
        );
    }

    /**
     * @param string $cell
     *
     * @return $this
     */
    public function setBottomRightCell(string $cell): Chart
    {
        $this->bottomRightCellRef = $cell;

        return $this;
    }

    /**
     * Get the cell address where the bottom right of the chart is fixed
     *
     * @return string
     */
    public function getBottomRightCell(): string
    {
        return $this->bottomRightCellRef;
    }

    /**
     * Set the offset position within the Bottom Right cell for the chart
     *
     * @param int|null $xOffset
     * @param int|null $yOffset
     *
     * @return $this
     */
    public function setBottomRightOffset(?int $xOffset = null, ?int $yOffset = null): Chart
    {
        if ($xOffset !== null) {
            $this->setBottomRightXOffset($xOffset);
        }
        if ($yOffset !== null) {
            $this->setBottomRightYOffset($yOffset);
        }

        return $this;
    }

    /**
     * Get the offset position within the Bottom Right cell for the chart
     *
     * @return integer[]
     */
    public function getBottomRightOffset(): array
    {
        return [
            'X' => $this->bottomRightXOffset,
            'Y' => $this->bottomRightYOffset
        ];
    }

    /**
     * @param int $xOffset
     *
     * @return $this
     */
    public function setBottomRightXOffset(int $xOffset): Chart
    {
        $this->bottomRightXOffset = $xOffset;

        return $this;
    }

    /**
     * @return int
     */
    public function getBottomRightXOffset(): int
    {
        return $this->bottomRightXOffset;
    }

    /**
     * @param int $yOffset
     *
     * @return $this
     */
    public function setBottomRightYOffset(int $yOffset): Chart
    {
        $this->bottomRightYOffset = $yOffset;

        return $this;
    }

    /**
     * @return int
     */
    public function getBottomRightYOffset(): int
    {
        return $this->bottomRightYOffset;
    }

    /**
     * @param Sheet $sheet
     *
     * @return $this
     */
    public function setSheet(Sheet $sheet): Chart
    {
        foreach ($this->plotArea->getPlotDataSeries() as $plotSeries) {
            if ($plotSeries) {
                $plotSeries->applyDataSourceSheet($sheet);
            }
        }
        $this->sheet = $sheet;

        return $this;
    }

    /**
     * @param array $colors
     *
     * @return $this
     */
    public function setChartColors(array $colors): Chart
    {
        $this->plotArea->setDefaultColors($colors);

        return $this;
    }

    /**
     * @param string $chartType
     *
     * @return $this
     */
    public function setChartType(string $chartType): Chart
    {
        if (!in_array($chartType, self::$charTypes)) {
            ExceptionChart::throwNew('Invalid chart type "' . $chartType . '"');
        }
        $this->chartType = $chartType;
        $this->plotArea->setChartType($chartType);

        return $this;
    }

    /**
     * Get the data series type(s) for a chart plot series
     *
     * @return array
     */
    public function getPlotChartTypes(): array
    {
        return $this->plotArea->getChartTypes();
    }

}