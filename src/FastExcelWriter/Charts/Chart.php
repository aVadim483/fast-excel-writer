<?php

namespace avadim\FastExcelWriter\Charts;

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Exceptions\Exception;
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
     * X-Axis Label
     *
     * @var Title|null
     */
    private $xAxisLabel = null;

    /**
     * Y-Axis Label
     *
     * @var Title|null
     */
    private $yAxisLabel = null;

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
     * Chart Asix Y as
     *
     * @var Axis|null
     */
    private ?Axis $yAxis;

    /**
     * Chart Asix X as
     *
     * @var Axis|null
     */
    private ?Axis $xAxis;

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
     * Create a new Chart
     */
    public function __construct($title, $plotArea, Legend $legend = null, $plotVisibleOnly = true, $displayBlanksAs = '0', $xAxisLabel = null, $yAxisLabel = null, Axis $xAxis = null, Axis $yAxis = null, GridLines $majorGridlines = null, GridLines $minorGridlines = null)
    {
        $this->setTitle($title);
        $this->legend = $legend;
        if ($xAxisLabel) {
            $this->xAxisLabel = is_string($xAxisLabel) ? (new Title($xAxisLabel)) : $xAxisLabel;
        }
        if ($yAxisLabel) {
            $this->yAxisLabel = is_string($yAxisLabel) ? (new Title($yAxisLabel)) : $yAxisLabel;
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
     * @param $chartType
     * @param $title
     * @param $dataSource
     *
     * @return $this
     */
    public static function make($chartType, $title = null, $dataSource = null): Chart
    {
        $plotLabels = [];
        $plotCategory = [];
        if ($dataSource instanceof PlotArea) {
            $plotArea = $dataSource;
            $chart = new static($title, $plotArea);
            $chart->setChartType($chartType);
        }
        else {
            $dataSeries = null;
            if ($dataSource instanceof DataSeries) {
                $dataSeries = $dataSource;
            }
            elseif ($dataSource instanceof DataSeriesValues) {
                $dataSeries = new DataSeries(
                    $chartType,		// plotType
                    [$dataSource],		// plotValues
                );
            }
            elseif (is_array($dataSource)) {
                $dataSeriesValues = [];
                $dataSeriesNames = [];
                foreach ($dataSource as $name => $data) {
                    if ($data instanceof DataSeriesValues) {
                        $dataSeriesValues[] = $data;
                    }
                    elseif (is_string($data)) {
                        $dimension = Excel::rangeDimension(str_replace('$', '', $data));
                        $dataSeriesValues[] = new DataSeriesValues('Number', $dimension['absAddress'], NULL, $dimension['cellCount']);

                    }
                    if (!is_int($name)) {
                        $dimension = Excel::rangeDimension(str_replace('$', '', $name));
                        $dataSeriesNames[] = new DataSeriesValues('String', $dimension['absAddress'], NULL, 1);
                    }
                    else {
                        $dataSeriesNames[] = null;
                    }
                }

                $dataSeries = new DataSeries(
                    $chartType,         // plotType
                    $dataSeriesValues,  // plotValues
                    $dataSeriesNames,   // plotLabels
                );
            }
            $plotArea = new PlotArea(NULL, [$dataSeries]);
            $chart = new static($title, $plotArea);
            $chart->setChartType($chartType);
        }

        return $chart;
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
        $this->title = is_string($title) ? (new Title($title)) : $title;

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

    public function setPlotArea($plotValues)
    {
        if ($plotValues instanceof PlotArea) {
            $this->plotArea = $plotValues;
        }
        elseif ($plotValues instanceof DataSeries) {
            $this->plotArea = new PlotArea(NULL, [$plotValues]);
        }
        elseif (is_array($plotValues)) {
            $plotValues = array_values($plotValues);
            if ($plotValues[0] instanceof DataSeries) {
                $this->plotArea = new PlotArea(NULL, $plotValues);
            }
        }
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
            $this->legend = new Legend($position, NULL, false);
        }
        else {
            $this->legend->setPosition($position);
        }

        return $this;
    }

    /**
     * @param $range
     *
     * @return $this
     */
    public function setDataSeriesTickLabels($range): Chart
    {
        $dataSeries = $this->plotArea->getPlotGroupByIndex(0);
        $dataSeries->setPlotCategories($range);

        return $this;
    }

    /**
     * Set X-Axis Label
     *
     * @param Title|string $label
     *
     * @return $this
     */
    public function setXAxisLabel($label): Chart
    {
        $this->xAxisLabel = is_string($label) ? (new Title($label)) : $label;

        return $this;
    }

    /**
     * Get X-Axis Label
     *
     * @return Title
     */
    public function getXAxisLabel(): ?Title
    {
        return $this->xAxisLabel;
    }

    /**
     * Set Y-Axis Label
     *
     * @param Title|string $label
     *
     * @return $this
     */
    public function setYAxisLabel($label): Chart
    {
        $this->yAxisLabel = is_string($label) ? (new Title($label)) : $label;

        return $this;
    }

    /**
     * Get Y-Axis Label
     *
     * @return Title
     */
    public function getYAxisLabel(): ?Title
    {
        return $this->yAxisLabel;
    }

    public function setDataSeriesNames($labels): Chart
    {
        $dataSeries = $this->plotArea->getPlotGroupByIndex(0);
        $dataSeries->setPlotLabels($labels);

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
    public function setTopLeftPosition(string $cell, int $xOffset = null, int $yOffset = null): Chart
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
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
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
    public function getTopLeftCell()
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
    public function setTopLeftCell(string $cell)
    {
        $this->topLeftCellRef = $cell;

        return $this;
    }

    /**
     * Set the offset position within the Top Left cell for the chart
     *
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return $this
     */
    public function setTopLeftOffset($xOffset = null, $yOffset = null)
    {
        if (!is_null($xOffset)) {
            $this->setTopLeftXOffset($xOffset);
        }
        if (!is_null($yOffset)) {
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

    public function setTopLeftXOffset($xOffset)
    {
        $this->topLeftXOffset = $xOffset;

        return $this;
    }

    public function getTopLeftXOffset(): int
    {
        return $this->topLeftXOffset;
    }

    public function setTopLeftYOffset($yOffset): Chart
    {
        $this->topLeftYOffset = $yOffset;

        return $this;
    }

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
    public function setBottomRightPosition(string $cell, ?int $xOffset = null, ?int $yOffset = null): Chart
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
        return array(
            'X' => $this->bottomRightXOffset,
            'Y' => $this->bottomRightYOffset
        );
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
        foreach ($this->plotArea->getPlotGroup() as $plotSeries) {
            $plotSeries->applyDataSourceSheet($sheet);
        }
        $this->sheet = $sheet;

        return $this;
    }

    /**
     * @param string $chartType
     *
     * @return $this
     */
    public function setChartType(string $chartType): Chart
    {
        if ($this->plotArea->getPlotSeriesCount() > 0) {
            $plotType = $chartType;
            if (substr($chartType, -8) === '_stacked') {
                $plotType = str_replace('_stacked', '', $plotType);
            }
            if ($plotType === self::TYPE_COLUMN) {
                $plotType = DataSeries::TYPE_BARCHART;
            }
            elseif (substr($plotType, -5) !== 'Chart') {
                $plotType .= 'Chart';
            }
            $plotSeries = $this->plotArea->getPlotGroupByIndex(0);
            $plotSeries->setPlotType($plotType);
            if ($chartType === self::TYPE_COLUMN || $chartType === self::TYPE_COLUMN_STACKED) {
                $plotSeries->setPlotDirection(DataSeries::DIRECTION_COL);
            }
            elseif (in_array($chartType, [self::TYPE_BAR, self::TYPE_BAR_STACKED])) {
                $plotSeries->setPlotDirection(DataSeries::DIRECTION_BAR);
            }
            if ($plotSeries->getPlotGrouping() === null) {
                if (substr($chartType, -8) === '_stacked') {
                    $plotSeries->setPlotGrouping(DataSeries::GROUPING_STACKED);
                }
                elseif (in_array($chartType, [self::TYPE_BAR])) {
                    $plotSeries->setPlotGrouping(DataSeries::GROUPING_CLUSTERED);
                }
                elseif (in_array($plotType, [DataSeries::TYPE_BARCHART, DataSeries::TYPE_BARCHART_3D, DataSeries::TYPE_LINECHART, DataSeries::TYPE_LINECHART_3D])) {
                    $plotSeries->setPlotGrouping(DataSeries::GROUPING_STANDARD);
                }
            }
        }

        return $this;
    }

    /**
     * Get the data series type(s) for a chart plot series
     *
     * @param PlotArea $plotArea
     *
     * @return array
     *
     * @throws Exception
     */
    private function getChartTypes(PlotArea $plotArea): array
    {
        $groupCount = $plotArea->getPlotGroupCount();

        if ($groupCount == 1) {
            $chartType = [$plotArea->getPlotGroupByIndex(0)->getPlotType()];
        }
        else {
            $chartTypes = [];
            for ($i = 0; $i < $groupCount; ++$i) {
                $chartTypes[] = $plotArea->getPlotGroupByIndex($i)->getPlotType();
            }
            $chartType = array_unique($chartTypes);
            if (count($chartTypes) == 0) {
                throw new Exception('Chart is not yet implemented');
            }
        }

        return $chartType;
    }

    /////////////////////////////////////////////////////////////////////////////////
    /// Write Chart to the file
    ///

    public function writeChart(FileWriter $fileWriter)
    {
        $relations = [];
        $fileWriter->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        $fileWriter->startElement('c:chartSpace');
        $fileWriter->writeAttribute('xmlns:c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
        $fileWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $fileWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        $fileWriter->startElement('c:date1904');
        $fileWriter->writeAttribute('val', 0);
        $fileWriter->endElement();
        $fileWriter->startElement('c:lang');
        $fileWriter->writeAttribute('val', "en-GB");
        $fileWriter->endElement();
        $fileWriter->startElement('c:roundedCorners');
        $fileWriter->writeAttribute('val', 0);
        $fileWriter->endElement();

        $fileWriter->startElement('mc:AlternateContent');
        $fileWriter->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');

        $fileWriter->startElement('mc:Choice');
        $fileWriter->writeAttribute('xmlns:c14', 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart');
        $fileWriter->writeAttribute('Requires', 'c14');

        $fileWriter->startElement('c14:style');
        $fileWriter->writeAttribute('val', '102');
        $fileWriter->endElement();
        $fileWriter->endElement();

        $fileWriter->startElement('mc:Fallback');
        $fileWriter->startElement('c:style');
        $fileWriter->writeAttribute('val', '2');
        $fileWriter->endElement();
        $fileWriter->endElement();

        $fileWriter->endElement();

        $fileWriter->startElement('c:chart');

        $title = $this->getTitle();
        $fileWriter->startElement('c:title');
        $fileWriter->startElement('c:tx');
        $fileWriter->startElement('c:rich');

        $fileWriter->writeElementAttr('a:bodyPr');

        $fileWriter->writeElementAttr('a:lstStyle');

        $fileWriter->startElement('a:p');

        $caption = $title ? $title->getCaption() : '';
        $fileWriter->startElement('<a:r>');
        $fileWriter->writeElementAttr('<a:t>' . $caption . '</a:t>');
        $fileWriter->endElement();
        $fileWriter->endElement();
        $fileWriter->endElement();
        $fileWriter->endElement();
        $fileWriter->writeElementAttr('<c:layout/>');
        $fileWriter->startElement('c:overlay');
        $fileWriter->writeAttribute('val', 0);
        $fileWriter->endElement();
        $fileWriter->endElement();

        $fileWriter->writeElementAttr('<c:autoTitleDeleted val="0"/>');

        $this->writePlotArea($fileWriter);

        if ($this->getLegend()) {
            $this->writeLegend($fileWriter, $this->getLegend());
        }

        $fileWriter->startElement('c:plotVisOnly');
        $fileWriter->writeAttribute('val', 1);
        $fileWriter->endElement();

        $fileWriter->startElement('c:dispBlanksAs');
        $fileWriter->writeAttribute('val', "gap");
        $fileWriter->endElement();

        $fileWriter->startElement('c:showDLblsOverMax');
        $fileWriter->writeAttribute('val', 0);
        $fileWriter->endElement();

        $fileWriter->endElement();

        //$this->writePrintSettings($fileWriter);

        $fileWriter->endElement();

        $fileWriter->flush(true);
    }

    /**
     * Write Chart Legend
     *
     * @param  FileWriter $fileWriter
     * @param  Legend|null $legend
     *
     * @throws  Exception
     */
    private function writeLegend(FileWriter $fileWriter, Legend $legend = null)
    {
        $fileWriter->startElement('c:legend');

        $fileWriter->startElement('c:legendPos');
        $fileWriter->writeAttribute('val', $legend->getPosition());
        $fileWriter->endElement();

        $this->writeLayout($fileWriter, $legend->getLayout());

        $fileWriter->startElement('c:overlay');
        $fileWriter->writeAttribute('val', ($legend->getOverlay()) ? '1' : '0');
        $fileWriter->endElement();

        $fileWriter->startElement('c:txPr');
        $fileWriter->startElement('a:bodyPr');
        $fileWriter->endElement();

        $fileWriter->startElement('a:lstStyle');
        $fileWriter->endElement();

        $fileWriter->startElement('a:p');
        $fileWriter->startElement('a:pPr');
        $fileWriter->writeAttribute('rtl', 0);

        $fileWriter->startElement('a:defRPr');
        $fileWriter->endElement();
        $fileWriter->endElement();

        $fileWriter->startElement('a:endParaRPr');
        $fileWriter->writeAttribute('lang', "en-US");
        $fileWriter->endElement();

        $fileWriter->endElement();
        $fileWriter->endElement();

        $fileWriter->endElement();
    }


    protected function writePlotArea(FileWriter $fileWriter)
    {
        $plotArea = $this->getPlotArea();
        $xAxisLabel = $this->getXAxisLabel();
        $yAxisLabel = $this->getYAxisLabel();
        $xAxis = $this->getChartAxisX();
        $yAxis = $this->getChartAxisY();
        $majorGridlines = $this->getMajorGridlines();
        $minorGridlines = $this->getMinorGridlines();

        $id1 = $id2 = 0;
        $this->_seriesIndex = 0;
        $fileWriter->startElement('c:plotArea');

        $layout = $plotArea->getLayout();

        $this->writeLayout($fileWriter, $layout);

        $chartTypes = $this->getChartTypes($plotArea);
        $catIsMultiLevelSeries = $valIsMultiLevelSeries = false;
        $plotGroupingType = '';

        foreach ($chartTypes as $chartType) {
            $fileWriter->startElement('c:' . $chartType);

            $groupCount = $plotArea->getPlotGroupCount();
            for ($i = 0; $i < $groupCount; ++$i) {
                $plotGroup = $plotArea->getPlotGroupByIndex($i);
                $groupType = $plotGroup->getPlotType();
                if ($groupType == $chartType) {
                    $plotStyle = $plotGroup->getPlotStyle();
                    if ($groupType === DataSeries::TYPE_RADARCHART) {
                        $fileWriter->startElement('c:radarStyle');
                        $fileWriter->writeAttribute('val', $plotStyle);
                        $fileWriter->endElement();
                    }
                    elseif ($groupType === DataSeries::TYPE_SCATTERCHART) {
                        $fileWriter->startElement('c:scatterStyle');
                        $fileWriter->writeAttribute('val', $plotStyle);
                        $fileWriter->endElement();
                    }

                    $this->writePlotGroup($fileWriter, $plotGroup, $chartType, $catIsMultiLevelSeries, $valIsMultiLevelSeries, $plotGroupingType);
                }
            }

            $this->writeDataLabels($fileWriter, $layout);

            if ($chartType === DataSeries::TYPE_LINECHART) {
                //    Line only, Line3D can't be smoothed

                $fileWriter->startElement('c:smooth');
                $fileWriter->writeAttribute('val', (int)$plotGroup->getSmoothLine());
                $fileWriter->endElement();
            } 
            elseif ($chartType === DataSeries::TYPE_BARCHART || $chartType === DataSeries::TYPE_BARCHART_3D) {
                $fileWriter->startElement('c:gapWidth');
                $fileWriter->writeAttribute('val', 150);
                $fileWriter->endElement();

                if ($plotGroupingType === 'percentStacked' || $plotGroupingType === 'stacked') {
                    $fileWriter->startElement('c:overlap');
                    $fileWriter->writeAttribute('val', 100);
                    $fileWriter->endElement();
                }
            }
            elseif ($chartType === DataSeries::TYPE_BUBBLECHART) {
                $fileWriter->startElement('c:bubbleScale');
                $fileWriter->writeAttribute('val', 25);
                $fileWriter->endElement();

                $fileWriter->startElement('c:showNegBubbles');
                $fileWriter->writeAttribute('val', 0);
                $fileWriter->endElement();
            }
            elseif ($chartType === DataSeries::TYPE_STOCKCHART) {
                $fileWriter->startElement('c:hiLowLines');
                $fileWriter->endElement();

                $fileWriter->startElement('c:upDownBars');

                $fileWriter->startElement('c:gapWidth');
                $fileWriter->writeAttribute('val', 300);
                $fileWriter->endElement();

                $fileWriter->startElement('c:upBars');
                $fileWriter->endElement();

                $fileWriter->startElement('c:downBars');
                $fileWriter->endElement();

                $fileWriter->endElement();
            }

            //    Generate 2 unique numbers to use for axId values
            //                    $id1 = $id2 = rand(10000000,99999999);
            //                    do {
            //                        $id2 = rand(10000000,99999999);
            //                    } while ($id1 == $id2);
            $id1 = '75091328';
            $id2 = '75089408';

            if (($chartType !== DataSeries::TYPE_PIECHART) && ($chartType !== DataSeries::TYPE_PIECHART_3D) && ($chartType !== DataSeries::TYPE_DONUTCHART)) {
                $fileWriter->startElement('c:axId');
                $fileWriter->writeAttribute('val', $id1);
                $fileWriter->endElement();
                $fileWriter->startElement('c:axId');
                $fileWriter->writeAttribute('val', $id2);
                $fileWriter->endElement();
            }
            else {
                $fileWriter->startElement('c:firstSliceAng');
                $fileWriter->writeAttribute('val', 0);
                $fileWriter->endElement();

                if ($chartType === DataSeries::TYPE_DONUTCHART) {
                    $fileWriter->startElement('c:holeSize');
                    $fileWriter->writeAttribute('val', 50);
                    $fileWriter->endElement();
                }
            }

            $fileWriter->endElement();

            if (($chartType !== DataSeries::TYPE_PIECHART) && ($chartType !== DataSeries::TYPE_PIECHART_3D) && ($chartType !== DataSeries::TYPE_DONUTCHART)) {
                if ($chartType === DataSeries::TYPE_BUBBLECHART) {
                    $this->writeValueAxis($fileWriter, $plotArea, $xAxisLabel, $chartType, $id1, $id2, $catIsMultiLevelSeries, $xAxis, $yAxis, $majorGridlines, $minorGridlines);
                }
                else {
                    $this->writeCategoryAxis($fileWriter, $plotArea, $xAxisLabel, $chartType, $id1, $id2, $catIsMultiLevelSeries, $xAxis, $yAxis);
                }

                $this->writeValueAxis($fileWriter, $plotArea, $yAxisLabel, $chartType, $id1, $id2, $valIsMultiLevelSeries, $xAxis, $yAxis, $majorGridlines, $minorGridlines);
            }
        }

        $fileWriter->endElement();
    }

    /**
     * @param Layout|null $layout
     * @param FileWriter $fileWriter
     * 
     * @return void
     */
    private function writeLayout(FileWriter $fileWriter, ?Layout $layout)
    {
        $fileWriter->startElement('c:layout');

        if ($layout) {
            $fileWriter->startElement('c:manualLayout');

            $layoutTarget = $layout->getLayoutTarget();
            if (!is_null($layoutTarget)) {
                $fileWriter->startElement('c:layoutTarget');
                $fileWriter->writeAttribute('val', $layoutTarget);
                $fileWriter->endElement();
            }

            $xMode = $layout->getXMode();
            if ($xMode) {
                $fileWriter->startElement('c:xMode');
                $fileWriter->writeAttribute('val', $xMode);
                $fileWriter->endElement();
            }

            $yMode = $layout->getYMode();
            if ($yMode) {
                $fileWriter->startElement('c:yMode');
                $fileWriter->writeAttribute('val', $yMode);
                $fileWriter->endElement();
            }

            $x = $layout->getXPosition();
            if ($x) {
                $fileWriter->startElement('c:x');
                $fileWriter->writeAttribute('val', $x);
                $fileWriter->endElement();
            }

            $y = $layout->getYPosition();
            if ($y) {
                $fileWriter->startElement('c:y');
                $fileWriter->writeAttribute('val', $y);
                $fileWriter->endElement();
            }

            $w = $layout->getWidth();
            if ($w) {
                $fileWriter->startElement('c:w');
                $fileWriter->writeAttribute('val', $w);
                $fileWriter->endElement();
            }

            $h = $layout->getHeight();
            if ($h) {
                $fileWriter->startElement('c:h');
                $fileWriter->writeAttribute('val', $h);
                $fileWriter->endElement();
            }

            $fileWriter->endElement();
        }

        $fileWriter->endElement();
    }

    /**
     * Write Plot Group (series of related plots)
     *
     * @param  FileWriter $fileWriter
     * @param  DataSeries $plotGroup
     * @param  string $groupType Type of plot for data series
     * @param  boolean &$catIsMultiLevelSeries Is category a multi-series category
     * @param  boolean &$valIsMultiLevelSeries Is value set a multi-series set
     * @param  string &$plotGroupingType Type of grouping for multi-series values
     */
    private function writePlotGroup(FileWriter $fileWriter, $plotGroup, $groupType, &$catIsMultiLevelSeries, &$valIsMultiLevelSeries, &$plotGroupingType)
    {
        if (is_null($plotGroup)) {
            return;
        }

        if (($groupType == DataSeries::TYPE_BARCHART) || ($groupType == DataSeries::TYPE_BARCHART_3D)) {
            $fileWriter->startElement('c:barDir');
            $fileWriter->writeAttribute('val', $plotGroup->getPlotDirection());
            $fileWriter->endElement();
        }

        if (!is_null($plotGroup->getPlotGrouping())) {
            $plotGroupingType = $plotGroup->getPlotGrouping();
            $fileWriter->startElement('c:grouping');
            $fileWriter->writeAttribute('val', $plotGroupingType);
            $fileWriter->endElement();
        }

        //    Get these details before the loop, because we can use the count to check for varyColors
        $plotSeriesOrder = $plotGroup->getPlotOrder();
        $plotSeriesCount = count($plotSeriesOrder);

        if (($groupType !== DataSeries::TYPE_RADARCHART) && ($groupType !== DataSeries::TYPE_STOCKCHART)) {
            if ($groupType !== DataSeries::TYPE_LINECHART) {
                if (($groupType == DataSeries::TYPE_PIECHART) || ($groupType == DataSeries::TYPE_PIECHART_3D) || ($groupType == DataSeries::TYPE_DONUTCHART) || ($plotSeriesCount > 1)) {
                    $fileWriter->startElement('c:varyColors');
                    $fileWriter->writeAttribute('val', 1);
                    $fileWriter->endElement();
                }
                else {
                    $fileWriter->startElement('c:varyColors');
                    $fileWriter->writeAttribute('val', 0);
                    $fileWriter->endElement();
                }
            }
        }

        foreach ($plotSeriesOrder as $plotSeriesIdx => $plotSeriesRef) {
            $fileWriter->startElement('c:ser');

            $fileWriter->startElement('c:idx');
            $fileWriter->writeAttribute('val', $this->_seriesIndex + $plotSeriesIdx);
            $fileWriter->endElement();

            $fileWriter->startElement('c:order');
            $fileWriter->writeAttribute('val', $this->_seriesIndex + $plotSeriesRef);
            $fileWriter->endElement();

            if (($groupType == DataSeries::TYPE_PIECHART) || ($groupType == DataSeries::TYPE_PIECHART_3D) || ($groupType == DataSeries::TYPE_DONUTCHART)) {
                $fileWriter->startElement('c:dPt');
                $fileWriter->startElement('c:idx');
                $fileWriter->writeAttribute('val', 3);
                $fileWriter->endElement();

                $fileWriter->startElement('c:bubble3D');
                $fileWriter->writeAttribute('val', 0);
                $fileWriter->endElement();

                $fileWriter->startElement('c:spPr');
                $fileWriter->startElement('a:solidFill');
                $fileWriter->startElement('a:srgbClr');
                $fileWriter->writeAttribute('val', 'FF9900');
                $fileWriter->endElement();
                $fileWriter->endElement();
                $fileWriter->endElement();
                $fileWriter->endElement();
            }

            //    Labels
            $plotSeriesLabel = $plotGroup->getPlotLabelByIndex($plotSeriesRef);
            if ($plotSeriesLabel && ($plotSeriesLabel->getPointCount() > 0)) {
                $fileWriter->startElement('c:tx');
                $fileWriter->startElement('c:strRef');
                $this->writePlotSeriesLabel($fileWriter, $plotSeriesLabel);
                $fileWriter->endElement();
                $fileWriter->endElement();
            }

            //    Formatting for the points
            if (($groupType == DataSeries::TYPE_LINECHART) || ($groupType == DataSeries::TYPE_STOCKCHART)) {
                $fileWriter->startElement('c:spPr');
                $fileWriter->startElement('a:ln');
                $fileWriter->writeAttribute('w', 12700);
                if ($groupType == DataSeries::TYPE_STOCKCHART) {
                    $fileWriter->startElement('a:noFill');
                    $fileWriter->endElement();
                }
                $fileWriter->endElement();
                $fileWriter->endElement();
            }
            else {
                /* custom colors of data series
                $fileWriter->startElement('c:spPr');
                $fileWriter->startElement('a:solidFill');
                $fileWriter->startElement('a:srgbClr');
                $fileWriter->writeAttribute('val', '777777');
                $fileWriter->endElement();
                $fileWriter->endElement();
                $fileWriter->endElement();
                */
            }

            $plotSeriesValues = $plotGroup->getPlotValuesByIndex($plotSeriesRef);
            if ($plotSeriesValues) {
                $plotSeriesMarker = $plotSeriesValues->getPointMarker();
                if ($plotSeriesMarker) {
                    $fileWriter->startElement('c:marker');
                    $fileWriter->startElement('c:symbol');
                    $fileWriter->writeAttribute('val', $plotSeriesMarker);
                    $fileWriter->endElement();

                    if ($plotSeriesMarker !== 'none') {
                        $fileWriter->startElement('c:size');
                        $fileWriter->writeAttribute('val', 3);
                        $fileWriter->endElement();
                    }

                    $fileWriter->endElement();
                }
            }

            if (($groupType === DataSeries::TYPE_BARCHART) || ($groupType === DataSeries::TYPE_BARCHART_3D) || ($groupType === DataSeries::TYPE_BUBBLECHART)) {
                $fileWriter->startElement('c:invertIfNegative');
                $fileWriter->writeAttribute('val', 0);
                $fileWriter->endElement();
            }

            //    Category Labels
            $plotSeriesCategory = $plotGroup->getPlotCategoryByIndex($plotSeriesRef);
            if ($plotSeriesCategory && ($plotSeriesCategory->getPointCount() > 0)) {
                $catIsMultiLevelSeries = $catIsMultiLevelSeries || $plotSeriesCategory->isMultiLevelSeries();

                if (($groupType == DataSeries::TYPE_PIECHART) || ($groupType == DataSeries::TYPE_PIECHART_3D) || ($groupType == DataSeries::TYPE_DONUTCHART)) {
                    if (!is_null($plotGroup->getPlotStyle())) {
                        $plotStyle = $plotGroup->getPlotStyle();
                        if ($plotStyle) {
                            $fileWriter->startElement('c:explosion');
                            $fileWriter->writeAttribute('val', 25);
                            $fileWriter->endElement();
                        }
                    }
                }

                if (($groupType === DataSeries::TYPE_BUBBLECHART) || ($groupType === DataSeries::TYPE_SCATTERCHART)) {
                    $fileWriter->startElement('c:xVal');
                } else {
                    $fileWriter->startElement('c:cat');
                }

                $this->writePlotSeriesValues($fileWriter, $plotSeriesCategory, $groupType, 'str');
                $fileWriter->endElement();
            }

            //    Values
            if ($plotSeriesValues) {
                $valIsMultiLevelSeries = $valIsMultiLevelSeries || $plotSeriesValues->isMultiLevelSeries();

                if (($groupType === DataSeries::TYPE_BUBBLECHART) || ($groupType === DataSeries::TYPE_SCATTERCHART)) {
                    $fileWriter->startElement('c:yVal');
                } else {
                    $fileWriter->startElement('c:val');
                }

                $this->writePlotSeriesValues($fileWriter, $plotSeriesValues, $groupType, 'num');
                $fileWriter->endElement();
            }

            if ($groupType === DataSeries::TYPE_BUBBLECHART) {
                $this->writeBubbles($fileWriter, $plotSeriesValues);
            }

            $fileWriter->endElement();
        }

        $this->_seriesIndex += $plotSeriesIdx + 1;
    }

    /**
     * Write Plot Series Label
     *
     * @param FileWriter $fileWriter
     * @param DataSeriesValues $plotSeriesLabel
     */
    private function writePlotSeriesLabel(FileWriter $fileWriter, DataSeriesValues $plotSeriesLabel)
    {
        $fileWriter->startElement('c:f');
        $fileWriter->writeRawData($plotSeriesLabel->getDataSource());
        $fileWriter->endElement();

        $fileWriter->startElement('c:strCache');
        $fileWriter->startElement('c:ptCount');
        $fileWriter->writeAttribute('val', $plotSeriesLabel->getPointCount());
        $fileWriter->endElement();

        foreach ($plotSeriesLabel->getDataValues() as $plotLabelKey => $plotLabelValue) {
            $fileWriter->startElement('c:pt');
            $fileWriter->writeAttribute('idx', $plotLabelKey);

            $fileWriter->startElement('c:v');
            $fileWriter->writeRawData($plotLabelValue);
            $fileWriter->endElement();
            $fileWriter->endElement();
        }
        $fileWriter->endElement();
    }

    /**
     * Write Plot Series Values
     *
     * @param FileWriter $fileWriter
     * @param  DataSeriesValues $plotSeriesValues
     * @param  string $groupType Type of plot for dataseries
     * @param  string $dataType Datatype of series values
     */
    private function writePlotSeriesValues($fileWriter, $plotSeriesValues, $groupType, $dataType)
    {
        if (is_null($plotSeriesValues)) {
            return;
        }

        if ($plotSeriesValues->isMultiLevelSeries()) {
            $levelCount = $plotSeriesValues->multiLevelCount();

            $fileWriter->startElement('c:multiLvlStrRef');

            $fileWriter->startElement('c:f');
            $fileWriter->writeRawData($plotSeriesValues->getDataSource());
            $fileWriter->endElement();

            $fileWriter->startElement('c:multiLvlStrCache');

            $fileWriter->startElement('c:ptCount');
            $fileWriter->writeAttribute('val', $plotSeriesValues->getPointCount());
            $fileWriter->endElement();

            for ($level = 0; $level < $levelCount; ++$level) {
                $fileWriter->startElement('c:lvl');

                foreach ($plotSeriesValues->getDataValues() as $plotSeriesKey => $plotSeriesValue) {
                    if (isset($plotSeriesValue[$level])) {
                        $fileWriter->startElement('c:pt');
                        $fileWriter->writeAttribute('idx', $plotSeriesKey);

                        $fileWriter->startElement('c:v');
                        $fileWriter->writeRawData($plotSeriesValue[$level]);
                        $fileWriter->endElement();
                        $fileWriter->endElement();
                    }
                }

                $fileWriter->endElement();
            }

            $fileWriter->endElement();

            $fileWriter->endElement();
        } else {
            $fileWriter->startElement('c:' . $dataType . 'Ref');

            $fileWriter->startElement('c:f');
            $fileWriter->writeRawData($plotSeriesValues->getDataSource());
            $fileWriter->endElement();

            $fileWriter->startElement('c:' . $dataType . 'Cache');

            if (($groupType != DataSeries::TYPE_PIECHART) && ($groupType != DataSeries::TYPE_PIECHART_3D) && ($groupType != DataSeries::TYPE_DONUTCHART)) {
                if (($plotSeriesValues->getFormatCode() !== null) && ($plotSeriesValues->getFormatCode() !== '')) {
                    $fileWriter->startElement('c:formatCode');
                    $fileWriter->writeRawData($plotSeriesValues->getFormatCode());
                    $fileWriter->endElement();
                }
            }

            $fileWriter->startElement('c:ptCount');
            $fileWriter->writeAttribute('val', $plotSeriesValues->getPointCount());
            $fileWriter->endElement();

            $dataValues = $plotSeriesValues->getDataValues();
            if (!empty($dataValues)) {
                if (is_array($dataValues)) {
                    foreach ($dataValues as $plotSeriesKey => $plotSeriesValue) {
                        $fileWriter->startElement('c:pt');
                        $fileWriter->writeAttribute('idx', $plotSeriesKey);

                        $fileWriter->startElement('c:v');
                        $fileWriter->writeRawData($plotSeriesValue);
                        $fileWriter->endElement();
                        $fileWriter->endElement();
                    }
                }
            }

            $fileWriter->endElement();

            $fileWriter->endElement();
        }
    }

    /**
     * Write Bubble Chart Details
     *
     * @param FileWriter $fileWriter XML Writer
     * @param DataSeriesValues $plotSeriesValues
     */
    private function writeBubbles(FileWriter $fileWriter, DataSeriesValues $plotSeriesValues)
    {
        if (is_null($plotSeriesValues)) {
            return;
        }

        $fileWriter->startElement('c:bubbleSize');
        $fileWriter->startElement('c:numLit');

        $fileWriter->startElement('c:formatCode');
        $fileWriter->writeRawData('General');
        $fileWriter->endElement();

        $fileWriter->startElement('c:ptCount');
        $fileWriter->writeAttribute('val', $plotSeriesValues->getPointCount());
        $fileWriter->endElement();

        $dataValues = $plotSeriesValues->getDataValues();
        if (!empty($dataValues)) {
            if (is_array($dataValues)) {
                foreach ($dataValues as $plotSeriesKey => $plotSeriesValue) {
                    $fileWriter->startElement('c:pt');
                    $fileWriter->writeAttribute('idx', $plotSeriesKey);
                    $fileWriter->startElement('c:v');
                    $fileWriter->writeRawData(1);
                    $fileWriter->endElement();
                    $fileWriter->endElement();
                }
            }
        }

        $fileWriter->endElement();
        $fileWriter->endElement();

        $fileWriter->startElement('c:bubble3D');
        $fileWriter->writeAttribute('val', 0);
        $fileWriter->endElement();
    }

    /**
     * Write Data Labels
     *
     * @param FileWriter $fileWriter
     * @param Layout|null $chartLayout Chart layout
     *
     * @throws  Exception
     */
    private function writeDataLabels($fileWriter, $chartLayout)
    {
        $fileWriter->startElement('c:dLbls');

        $fileWriter->startElement('c:showLegendKey');
        $showLegendKey = !$chartLayout ? 0 : $chartLayout->getShowLegendKey();
        $fileWriter->writeAttribute('val', $showLegendKey ? 1 : 0);
        $fileWriter->endElement();

        $fileWriter->startElement('c:showVal');
        $showVal = !$chartLayout ? 0 : $chartLayout->getShowVal();
        $fileWriter->writeAttribute('val', $showVal ? 1 : 0);
        $fileWriter->endElement();

        $fileWriter->startElement('c:showCatName');
        $showCatName = !$chartLayout ? 0 : $chartLayout->getShowCatName();
        $fileWriter->writeAttribute('val', $showCatName ? 1 : 0);
        $fileWriter->endElement();

        $fileWriter->startElement('c:showSerName');
        $showSerName = !$chartLayout ? 0 : $chartLayout->getShowSerName();
        $fileWriter->writeAttribute('val', $showSerName ? 1 : 0);
        $fileWriter->endElement();

        $fileWriter->startElement('c:showPercent');
        $showPercent = !$chartLayout ? 0 : $chartLayout->getShowPercent();
        $fileWriter->writeAttribute('val', $showPercent ? 1 : 0);
        $fileWriter->endElement();

        $fileWriter->startElement('c:showBubbleSize');
        $showBubbleSize = !$chartLayout ? 0 : $chartLayout->getShowBubbleSize();
        $fileWriter->writeAttribute('val', $showBubbleSize ? 1 : 0);
        $fileWriter->endElement();

        $fileWriter->startElement('c:showLeaderLines');
        $showLeaderLines = !$chartLayout ? 1 : $chartLayout->getShowLeaderLines();
        $fileWriter->writeAttribute('val', $showLeaderLines ? 1 : 0);
        $fileWriter->endElement();

        $fileWriter->endElement();
    }

    /**
     * Write Category Axis
     *
     * @param FileWriter $fileWriter
     * @param PlotArea $plotArea
     * @param Title $xAxisLabel
     * @param string $groupType Chart type
     * @param string $id1
     * @param string $id2
     * @param boolean $isMultiLevelSeries
     * @param $xAxis
     * @param $yAxis
     */
    private function writeCategoryAxis($fileWriter, PlotArea $plotArea, $xAxisLabel, $groupType, $id1, $id2, $isMultiLevelSeries, $xAxis, $yAxis)
    {
        $fileWriter->startElement('c:catAx');

        if ($id1 > 0) {
            $fileWriter->startElement('c:axId');
            $fileWriter->writeAttribute('val', $id1);
            $fileWriter->endElement();
        }

        $fileWriter->startElement('c:scaling');
        $fileWriter->startElement('c:orientation');
        $fileWriter->writeAttribute('val', $yAxis->getAxisOptionsProperty('orientation'));
        $fileWriter->endElement();
        $fileWriter->endElement();

        $fileWriter->startElement('c:delete');
        $fileWriter->writeAttribute('val', 0);
        $fileWriter->endElement();

        $fileWriter->startElement('c:axPos');
        $fileWriter->writeAttribute('val', "b");
        $fileWriter->endElement();

        if (!is_null($xAxisLabel)) {
            $fileWriter->startElement('c:title');
            $fileWriter->startElement('c:tx');
            $fileWriter->startElement('c:rich');

            $fileWriter->startElement('a:bodyPr');
            $fileWriter->endElement();

            $fileWriter->startElement('a:lstStyle');
            $fileWriter->endElement();

            $fileWriter->startElement('a:p');
            $fileWriter->startElement('a:r');

            $caption = $xAxisLabel->getCaption();
            $fileWriter->startElement('a:t');
            //                                        $fileWriter->writeAttribute('xml:space', 'preserve');
            $fileWriter->writeRawData(Writer::xmlSpecialChars($caption));
            $fileWriter->endElement();

            $fileWriter->endElement();
            $fileWriter->endElement();
            $fileWriter->endElement();
            $fileWriter->endElement();

            $layout = $xAxisLabel->getLayout();
            $this->writeLayout($fileWriter, $layout);

            $fileWriter->startElement('c:overlay');
            $fileWriter->writeAttribute('val', 0);
            $fileWriter->endElement();

            $fileWriter->endElement();
        }

        $fileWriter->startElement('c:numFmt');
        $fileWriter->writeAttribute('formatCode', $yAxis->getAxisNumberFormat());
        $fileWriter->writeAttribute('sourceLinked', $yAxis->getAxisNumberSourceLinked());
        $fileWriter->endElement();

        $fileWriter->startElement('c:majorTickMark');
        $fileWriter->writeAttribute('val', $yAxis->getAxisOptionsProperty('major_tick_mark'));
        $fileWriter->endElement();

        $fileWriter->startElement('c:minorTickMark');
        $fileWriter->writeAttribute('val', $yAxis->getAxisOptionsProperty('minor_tick_mark'));
        $fileWriter->endElement();

        $fileWriter->startElement('c:tickLblPos');
        $fileWriter->writeAttribute('val', $yAxis->getAxisOptionsProperty('axis_labels'));
        $fileWriter->endElement();

        if ($id2 > 0) {
            $fileWriter->startElement('c:crossAx');
            $fileWriter->writeAttribute('val', $id2);
            $fileWriter->endElement();

            $fileWriter->startElement('c:crosses');
            $fileWriter->writeAttribute('val', $yAxis->getAxisOptionsProperty('horizontal_crosses'));
            $fileWriter->endElement();
        }

        $fileWriter->startElement('c:auto');
        $fileWriter->writeAttribute('val', 1);
        $fileWriter->endElement();

        $fileWriter->startElement('c:lblAlgn');
        $fileWriter->writeAttribute('val', "ctr");
        $fileWriter->endElement();

        $fileWriter->startElement('c:lblOffset');
        $fileWriter->writeAttribute('val', 100);
        $fileWriter->endElement();

        if ($isMultiLevelSeries) {
            $fileWriter->startElement('c:noMultiLvlLbl');
            $fileWriter->writeAttribute('val', 0);
            $fileWriter->endElement();
        }
        $fileWriter->endElement();
    }

    /**
     * Write Value Axis
     *
     * @param FileWriter $fileWriter
     * @param PlotArea $plotArea
     * @param Title $yAxisLabel
     * @param string $groupType Chart type
     * @param string $id1
     * @param string $id2
     * @param boolean $isMultiLevelSeries
     * @param $xAxis
     * @param $yAxis
     * @param $majorGridlines
     * @param $minorGridlines
     */
    private function writeValueAxis($fileWriter, PlotArea $plotArea, $yAxisLabel, $groupType, $id1, $id2, $isMultiLevelSeries, $xAxis, $yAxis, $majorGridlines, $minorGridlines)
    {
        $fileWriter->startElement('c:valAx');

        if ($id2 > 0) {
            $fileWriter->startElement('c:axId');
            $fileWriter->writeAttribute('val', $id2);
            $fileWriter->endElement();
        }

        $fileWriter->startElement('c:scaling');

        if (!is_null($xAxis->getAxisOptionsProperty('maximum'))) {
            $fileWriter->startElement('c:max');
            $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('maximum'));
            $fileWriter->endElement();
        }

        if (!is_null($xAxis->getAxisOptionsProperty('minimum'))) {
            $fileWriter->startElement('c:min');
            $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('minimum'));
            $fileWriter->endElement();
        }

        $fileWriter->startElement('c:orientation');
        $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('orientation'));


        $fileWriter->endElement();
        $fileWriter->endElement();

        $fileWriter->startElement('c:delete');
        $fileWriter->writeAttribute('val', 0);
        $fileWriter->endElement();

        $fileWriter->startElement('c:axPos');
        $fileWriter->writeAttribute('val', "l");
        $fileWriter->endElement();

        $fileWriter->startElement('c:majorGridlines');
        $fileWriter->startElement('c:spPr');

        if (!is_null($majorGridlines->getLineColorProperty('value'))) {
            $fileWriter->startElement('a:ln');
            $fileWriter->writeAttribute('w', $majorGridlines->getLineStyleProperty('width'));
            $fileWriter->startElement('a:solidFill');
            $fileWriter->startElement("a:{$majorGridlines->getLineColorProperty('type')}");
            $fileWriter->writeAttribute('val', $majorGridlines->getLineColorProperty('value'));
            $fileWriter->startElement('a:alpha');
            $fileWriter->writeAttribute('val', $majorGridlines->getLineColorProperty('alpha'));
            $fileWriter->endElement(); //end alpha
            $fileWriter->endElement(); //end srgbClr
            $fileWriter->endElement(); //end solidFill

            $fileWriter->startElement('a:prstDash');
            $fileWriter->writeAttribute('val', $majorGridlines->getLineStyleProperty('dash'));
            $fileWriter->endElement();

            if ($majorGridlines->getLineStyleProperty('join') == 'miter') {
                $fileWriter->startElement('a:miter');
                $fileWriter->writeAttribute('lim', '800000');
                $fileWriter->endElement();
            }
            else {
                $fileWriter->startElement('a:bevel');
                $fileWriter->endElement();
            }

            if (!is_null($majorGridlines->getLineStyleProperty(array('arrow', 'head', 'type')))) {
                $fileWriter->startElement('a:headEnd');
                $fileWriter->writeAttribute('type', $majorGridlines->getLineStyleProperty(array('arrow', 'head', 'type')));
                $fileWriter->writeAttribute('w', $majorGridlines->getLineStyleArrowParameters('head', 'w'));
                $fileWriter->writeAttribute('len', $majorGridlines->getLineStyleArrowParameters('head', 'len'));
                $fileWriter->endElement();
            }

            if (!is_null($majorGridlines->getLineStyleProperty(array('arrow', 'end', 'type')))) {
                $fileWriter->startElement('a:tailEnd');
                $fileWriter->writeAttribute('type', $majorGridlines->getLineStyleProperty(array('arrow', 'end', 'type')));
                $fileWriter->writeAttribute('w', $majorGridlines->getLineStyleArrowParameters('end', 'w'));
                $fileWriter->writeAttribute('len', $majorGridlines->getLineStyleArrowParameters('end', 'len'));
                $fileWriter->endElement();
            }
            $fileWriter->endElement(); //end ln
        }
        $fileWriter->startElement('a:effectLst');

        if (!is_null($majorGridlines->getGlowSize())) {
            $fileWriter->startElement('a:glow');
            $fileWriter->writeAttribute('rad', $majorGridlines->getGlowSize());
            $fileWriter->startElement("a:{$majorGridlines->getGlowColor('type')}");
            $fileWriter->writeAttribute('val', $majorGridlines->getGlowColor('value'));
            $fileWriter->startElement('a:alpha');
            $fileWriter->writeAttribute('val', $majorGridlines->getGlowColor('alpha'));
            $fileWriter->endElement(); //end alpha
            $fileWriter->endElement(); //end schemeClr
            $fileWriter->endElement(); //end glow
        }

        if (!is_null($majorGridlines->getShadowProperty('presets'))) {
            $fileWriter->startElement("a:{$majorGridlines->getShadowProperty('effect')}");
            if (!is_null($majorGridlines->getShadowProperty('blur'))) {
                $fileWriter->writeAttribute('blurRad', $majorGridlines->getShadowProperty('blur'));
            }
            if (!is_null($majorGridlines->getShadowProperty('distance'))) {
                $fileWriter->writeAttribute('dist', $majorGridlines->getShadowProperty('distance'));
            }
            if (!is_null($majorGridlines->getShadowProperty('direction'))) {
                $fileWriter->writeAttribute('dir', $majorGridlines->getShadowProperty('direction'));
            }
            if (!is_null($majorGridlines->getShadowProperty('algn'))) {
                $fileWriter->writeAttribute('algn', $majorGridlines->getShadowProperty('algn'));
            }
            if (!is_null($majorGridlines->getShadowProperty(array('size', 'sx')))) {
                $fileWriter->writeAttribute('sx', $majorGridlines->getShadowProperty(array('size', 'sx')));
            }
            if (!is_null($majorGridlines->getShadowProperty(array('size', 'sy')))) {
                $fileWriter->writeAttribute('sy', $majorGridlines->getShadowProperty(array('size', 'sy')));
            }
            if (!is_null($majorGridlines->getShadowProperty(array('size', 'kx')))) {
                $fileWriter->writeAttribute('kx', $majorGridlines->getShadowProperty(array('size', 'kx')));
            }
            if (!is_null($majorGridlines->getShadowProperty('rotWithShape'))) {
                $fileWriter->writeAttribute('rotWithShape', $majorGridlines->getShadowProperty('rotWithShape'));
            }
            $fileWriter->startElement("a:{$majorGridlines->getShadowProperty(array('color', 'type'))}");
            $fileWriter->writeAttribute('val', $majorGridlines->getShadowProperty(array('color', 'value')));

            $fileWriter->startElement('a:alpha');
            $fileWriter->writeAttribute('val', $majorGridlines->getShadowProperty(array('color', 'alpha')));
            $fileWriter->endElement(); //end alpha

            $fileWriter->endElement(); //end color:type
            $fileWriter->endElement(); //end shadow
        }

        if (!is_null($majorGridlines->getSoftEdgesSize())) {
            $fileWriter->startElement('a:softEdge');
            $fileWriter->writeAttribute('rad', $majorGridlines->getSoftEdgesSize());
            $fileWriter->endElement(); //end softEdge
        }

        $fileWriter->endElement(); //end effectLst
        $fileWriter->endElement(); //end spPr
        $fileWriter->endElement(); //end majorGridLines

        if ($minorGridlines->getObjectState()) {
            $fileWriter->startElement('c:minorGridlines');
            $fileWriter->startElement('c:spPr');

            if (!is_null($minorGridlines->getLineColorProperty('value'))) {
                $fileWriter->startElement('a:ln');
                $fileWriter->writeAttribute('w', $minorGridlines->getLineStyleProperty('width'));
                $fileWriter->startElement('a:solidFill');
                $fileWriter->startElement("a:{$minorGridlines->getLineColorProperty('type')}");
                $fileWriter->writeAttribute('val', $minorGridlines->getLineColorProperty('value'));
                $fileWriter->startElement('a:alpha');
                $fileWriter->writeAttribute('val', $minorGridlines->getLineColorProperty('alpha'));
                $fileWriter->endElement(); //end alpha
                $fileWriter->endElement(); //end srgbClr
                $fileWriter->endElement(); //end solidFill

                $fileWriter->startElement('a:prstDash');
                $fileWriter->writeAttribute('val', $minorGridlines->getLineStyleProperty('dash'));
                $fileWriter->endElement();

                if ($minorGridlines->getLineStyleProperty('join') == 'miter') {
                    $fileWriter->startElement('a:miter');
                    $fileWriter->writeAttribute('lim', '800000');
                    $fileWriter->endElement();
                }
                else {
                    $fileWriter->startElement('a:bevel');
                    $fileWriter->endElement();
                }

                if (!is_null($minorGridlines->getLineStyleProperty(array('arrow', 'head', 'type')))) {
                    $fileWriter->startElement('a:headEnd');
                    $fileWriter->writeAttribute('type', $minorGridlines->getLineStyleProperty(array('arrow', 'head', 'type')));
                    $fileWriter->writeAttribute('w', $minorGridlines->getLineStyleArrowParameters('head', 'w'));
                    $fileWriter->writeAttribute('len', $minorGridlines->getLineStyleArrowParameters('head', 'len'));
                    $fileWriter->endElement();
                }

                if (!is_null($minorGridlines->getLineStyleProperty(array('arrow', 'end', 'type')))) {
                    $fileWriter->startElement('a:tailEnd');
                    $fileWriter->writeAttribute('type', $minorGridlines->getLineStyleProperty(array('arrow', 'end', 'type')));
                    $fileWriter->writeAttribute('w', $minorGridlines->getLineStyleArrowParameters('end', 'w'));
                    $fileWriter->writeAttribute('len', $minorGridlines->getLineStyleArrowParameters('end', 'len'));
                    $fileWriter->endElement();
                }
                $fileWriter->endElement(); //end ln
            }

            $fileWriter->startElement('a:effectLst');

            if (!is_null($minorGridlines->getGlowSize())) {
                $fileWriter->startElement('a:glow');
                $fileWriter->writeAttribute('rad', $minorGridlines->getGlowSize());
                $fileWriter->startElement("a:{$minorGridlines->getGlowColor('type')}");
                $fileWriter->writeAttribute('val', $minorGridlines->getGlowColor('value'));
                $fileWriter->startElement('a:alpha');
                $fileWriter->writeAttribute('val', $minorGridlines->getGlowColor('alpha'));
                $fileWriter->endElement(); //end alpha
                $fileWriter->endElement(); //end schemeClr
                $fileWriter->endElement(); //end glow
            }

            if (!is_null($minorGridlines->getShadowProperty('presets'))) {
                $fileWriter->startElement("a:{$minorGridlines->getShadowProperty('effect')}");
                if (!is_null($minorGridlines->getShadowProperty('blur'))) {
                    $fileWriter->writeAttribute('blurRad', $minorGridlines->getShadowProperty('blur'));
                }
                if (!is_null($minorGridlines->getShadowProperty('distance'))) {
                    $fileWriter->writeAttribute('dist', $minorGridlines->getShadowProperty('distance'));
                }
                if (!is_null($minorGridlines->getShadowProperty('direction'))) {
                    $fileWriter->writeAttribute('dir', $minorGridlines->getShadowProperty('direction'));
                }
                if (!is_null($minorGridlines->getShadowProperty('algn'))) {
                    $fileWriter->writeAttribute('algn', $minorGridlines->getShadowProperty('algn'));
                }
                if (!is_null($minorGridlines->getShadowProperty(array('size', 'sx')))) {
                    $fileWriter->writeAttribute('sx', $minorGridlines->getShadowProperty(array('size', 'sx')));
                }
                if (!is_null($minorGridlines->getShadowProperty(array('size', 'sy')))) {
                    $fileWriter->writeAttribute('sy', $minorGridlines->getShadowProperty(array('size', 'sy')));
                }
                if (!is_null($minorGridlines->getShadowProperty(array('size', 'kx')))) {
                    $fileWriter->writeAttribute('kx', $minorGridlines->getShadowProperty(array('size', 'kx')));
                }
                if (!is_null($minorGridlines->getShadowProperty('rotWithShape'))) {
                    $fileWriter->writeAttribute('rotWithShape', $minorGridlines->getShadowProperty('rotWithShape'));
                }
                $fileWriter->startElement("a:{$minorGridlines->getShadowProperty(array('color', 'type'))}");
                $fileWriter->writeAttribute('val', $minorGridlines->getShadowProperty(array('color', 'value')));
                $fileWriter->startElement('a:alpha');
                $fileWriter->writeAttribute('val', $minorGridlines->getShadowProperty(array('color', 'alpha')));
                $fileWriter->endElement(); //end alpha
                $fileWriter->endElement(); //end color:type
                $fileWriter->endElement(); //end shadow
            }

            if (!is_null($minorGridlines->getSoftEdgesSize())) {
                $fileWriter->startElement('a:softEdge');
                $fileWriter->writeAttribute('rad', $minorGridlines->getSoftEdgesSize());
                $fileWriter->endElement(); //end softEdge
            }

            $fileWriter->endElement(); //end effectLst
            $fileWriter->endElement(); //end spPr
            $fileWriter->endElement(); //end minorGridLines
        }

        if (!is_null($yAxisLabel)) {
            $fileWriter->startElement('c:title');
            $fileWriter->startElement('c:tx');
            $fileWriter->startElement('c:rich');

            $fileWriter->startElement('a:bodyPr');
            $fileWriter->endElement();

            $fileWriter->startElement('a:lstStyle');
            $fileWriter->endElement();

            $fileWriter->startElement('a:p');
            $fileWriter->startElement('a:r');

            $caption = $yAxisLabel->getCaption();
            if (is_array($caption)) {
                $caption = $caption[0];
            }

            $fileWriter->startElement('a:t');
            //                                        $fileWriter->writeAttribute('xml:space', 'preserve');
            $fileWriter->writeRawData(Writer::xmlSpecialChars($caption));
            $fileWriter->endElement();

            $fileWriter->endElement();
            $fileWriter->endElement();
            $fileWriter->endElement();
            $fileWriter->endElement();

            if ($groupType !== DataSeries::TYPE_BUBBLECHART) {
                $layout = $yAxisLabel->getLayout();
                $this->writeLayout($fileWriter, $layout);
            }

            $fileWriter->startElement('c:overlay');
            $fileWriter->writeAttribute('val', 0);
            $fileWriter->endElement();

            $fileWriter->endElement();
        }

        $fileWriter->startElement('c:numFmt');
        $fileWriter->writeAttribute('formatCode', $xAxis->getAxisNumberFormat());
        $fileWriter->writeAttribute('sourceLinked', $xAxis->getAxisNumberSourceLinked());
        $fileWriter->endElement();

        $fileWriter->startElement('c:majorTickMark');
        $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('major_tick_mark'));
        $fileWriter->endElement();

        $fileWriter->startElement('c:minorTickMark');
        $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('minor_tick_mark'));
        $fileWriter->endElement();

        $fileWriter->startElement('c:tickLblPos');
        $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('axis_labels'));
        $fileWriter->endElement();

        $fileWriter->startElement('c:spPr');

        if (!is_null($xAxis->getFillProperty('value'))) {
            $fileWriter->startElement('a:solidFill');
            $fileWriter->startElement("a:" . $xAxis->getFillProperty('type'));
            $fileWriter->writeAttribute('val', $xAxis->getFillProperty('value'));
            $fileWriter->startElement('a:alpha');
            $fileWriter->writeAttribute('val', $xAxis->getFillProperty('alpha'));
            $fileWriter->endElement();
            $fileWriter->endElement();
            $fileWriter->endElement();
        }

        $fileWriter->startElement('a:ln');

        $fileWriter->writeAttribute('w', $xAxis->getLineStyleProperty('width'));
        $fileWriter->writeAttribute('cap', $xAxis->getLineStyleProperty('cap'));
        $fileWriter->writeAttribute('cmpd', $xAxis->getLineStyleProperty('compound'));

        if (!is_null($xAxis->getLineProperty('value'))) {
            $fileWriter->startElement('a:solidFill');
            $fileWriter->startElement("a:" . $xAxis->getLineProperty('type'));
            $fileWriter->writeAttribute('val', $xAxis->getLineProperty('value'));
            $fileWriter->startElement('a:alpha');
            $fileWriter->writeAttribute('val', $xAxis->getLineProperty('alpha'));
            $fileWriter->endElement();
            $fileWriter->endElement();
            $fileWriter->endElement();
        }

        $fileWriter->startElement('a:prstDash');
        $fileWriter->writeAttribute('val', $xAxis->getLineStyleProperty('dash'));
        $fileWriter->endElement();

        if ($xAxis->getLineStyleProperty('join') == 'miter') {
            $fileWriter->startElement('a:miter');
            $fileWriter->writeAttribute('lim', '800000');
            $fileWriter->endElement();
        }
        else {
            $fileWriter->startElement('a:bevel');
            $fileWriter->endElement();
        }

        if (!is_null($xAxis->getLineStyleProperty(array('arrow', 'head', 'type')))) {
            $fileWriter->startElement('a:headEnd');
            $fileWriter->writeAttribute('type', $xAxis->getLineStyleProperty(array('arrow', 'head', 'type')));
            $fileWriter->writeAttribute('w', $xAxis->getLineStyleArrowWidth('head'));
            $fileWriter->writeAttribute('len', $xAxis->getLineStyleArrowLength('head'));
            $fileWriter->endElement();
        }

        if (!is_null($xAxis->getLineStyleProperty(['arrow', 'end', 'type']))) {
            $fileWriter->startElement('a:tailEnd');
            $fileWriter->writeAttribute('type', $xAxis->getLineStyleProperty(['arrow', 'end', 'type']));
            $fileWriter->writeAttribute('w', $xAxis->getLineStyleArrowWidth('end'));
            $fileWriter->writeAttribute('len', $xAxis->getLineStyleArrowLength('end'));
            $fileWriter->endElement();
        }

        $fileWriter->endElement();

        $fileWriter->startElement('a:effectLst');

        if (!is_null($xAxis->getGlowProperty('size'))) {
            $fileWriter->startElement('a:glow');
            $fileWriter->writeAttribute('rad', $xAxis->getGlowProperty('size'));
            $fileWriter->startElement("a:{$xAxis->getGlowProperty(['color','type'])}");
            $fileWriter->writeAttribute('val', $xAxis->getGlowProperty(['color','value']));
            $fileWriter->startElement('a:alpha');
            $fileWriter->writeAttribute('val', $xAxis->getGlowProperty(['color','alpha']));
            $fileWriter->endElement();
            $fileWriter->endElement();
            $fileWriter->endElement();
        }

        if (!is_null($xAxis->getShadowProperty('presets'))) {
            $fileWriter->startElement("a:{$xAxis->getShadowProperty('effect')}");

            if (!is_null($xAxis->getShadowProperty('blur'))) {
                $fileWriter->writeAttribute('blurRad', $xAxis->getShadowProperty('blur'));
            }
            if (!is_null($xAxis->getShadowProperty('distance'))) {
                $fileWriter->writeAttribute('dist', $xAxis->getShadowProperty('distance'));
            }
            if (!is_null($xAxis->getShadowProperty('direction'))) {
                $fileWriter->writeAttribute('dir', $xAxis->getShadowProperty('direction'));
            }
            if (!is_null($xAxis->getShadowProperty('algn'))) {
                $fileWriter->writeAttribute('algn', $xAxis->getShadowProperty('algn'));
            }
            if (!is_null($xAxis->getShadowProperty(array('size','sx')))) {
                $fileWriter->writeAttribute('sx', $xAxis->getShadowProperty(array('size','sx')));
            }
            if (!is_null($xAxis->getShadowProperty(array('size','sy')))) {
                $fileWriter->writeAttribute('sy', $xAxis->getShadowProperty(array('size','sy')));
            }
            if (!is_null($xAxis->getShadowProperty(array('size','kx')))) {
                $fileWriter->writeAttribute('kx', $xAxis->getShadowProperty(array('size','kx')));
            }
            if (!is_null($xAxis->getShadowProperty('rotWithShape'))) {
                $fileWriter->writeAttribute('rotWithShape', $xAxis->getShadowProperty('rotWithShape'));
            }

            $fileWriter->startElement("a:{$xAxis->getShadowProperty(array('color','type'))}");
            $fileWriter->writeAttribute('val', $xAxis->getShadowProperty(array('color','value')));
            $fileWriter->startElement('a:alpha');
            $fileWriter->writeAttribute('val', $xAxis->getShadowProperty(array('color','alpha')));
            $fileWriter->endElement();
            $fileWriter->endElement();

            $fileWriter->endElement();
        }

        if (!is_null($xAxis->getSoftEdgesSize())) {
            $fileWriter->startElement('a:softEdge');
            $fileWriter->writeAttribute('rad', $xAxis->getSoftEdgesSize());
            $fileWriter->endElement();
        }

        $fileWriter->endElement(); //effectList
        $fileWriter->endElement(); //end spPr

        if ($id1 > 0) {
            $fileWriter->startElement('c:crossAx');
            $fileWriter->writeAttribute('val', $id2);
            $fileWriter->endElement();

            if (!is_null($xAxis->getAxisOptionsProperty('horizontal_crosses_value'))) {
                $fileWriter->startElement('c:crossesAt');
                $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('horizontal_crosses_value'));
                $fileWriter->endElement();
            } else {
                $fileWriter->startElement('c:crosses');
                $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('horizontal_crosses'));
                $fileWriter->endElement();
            }

            $fileWriter->startElement('c:crossBetween');
            $fileWriter->writeAttribute('val', "midCat");
            $fileWriter->endElement();

            if (!is_null($xAxis->getAxisOptionsProperty('major_unit'))) {
                $fileWriter->startElement('c:majorUnit');
                $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('major_unit'));
                $fileWriter->endElement();
            }

            if (!is_null($xAxis->getAxisOptionsProperty('minor_unit'))) {
                $fileWriter->startElement('c:minorUnit');
                $fileWriter->writeAttribute('val', $xAxis->getAxisOptionsProperty('minor_unit'));
                $fileWriter->endElement();
            }
        }

        if ($isMultiLevelSeries) {
            if ($groupType !== DataSeries::TYPE_BUBBLECHART) {
                $fileWriter->startElement('c:noMultiLvlLbl');
                $fileWriter->writeAttribute('val', 0);
                $fileWriter->endElement();
            }
        }

        $fileWriter->endElement();
    }


}