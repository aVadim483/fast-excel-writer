# Class \avadim\FastExcelWriter\Charts\Chart

---

* [__construct()](#__construct) -- Chart constructor
* [make()](#make)
* [addDataSeriesSet()](#adddataseriesset)
* [addDataSeriesType()](#adddataseriestype)
* [addDataSeriesType2()](#adddataseriestype2)
* [addDataSeriesValues()](#adddataseriesvalues)
* [getBottomRightCell()](#getbottomrightcell) -- Get the cell address where the bottom right of the chart is fixed
* [setBottomRightCell()](#setbottomrightcell)
* [getBottomRightOffset()](#getbottomrightoffset) -- Get the offset position within the Bottom Right cell for the chart
* [setBottomRightOffset()](#setbottomrightoffset) -- Set the offset position within the Bottom Right cell for the chart
* [getBottomRightPosition()](#getbottomrightposition) -- Get the bottom right position of the chart
* [getBottomRightXOffset()](#getbottomrightxoffset)
* [setBottomRightXOffset()](#setbottomrightxoffset)
* [getBottomRightYOffset()](#getbottomrightyoffset)
* [setBottomRightYOffset()](#setbottomrightyoffset)
* [setCategoryAxis()](#setcategoryaxis)
* [setCategoryAxisLabels()](#setcategoryaxislabels) -- Set Category Axis Labels
* [getCategoryAxisTitle()](#getcategoryaxistitle) -- Get Category Axis Title
* [setCategoryAxisTitle()](#setcategoryaxistitle) -- Set Category Axis Title
* [getChartAxisX()](#getchartaxisx) -- Get xAxis
* [getChartAxisY()](#getchartaxisy) -- Get yAxis
* [getChartAxisY2()](#getchartaxisy2) -- Get yAxis
* [setChartColors()](#setchartcolors)
* [setChartType()](#setcharttype)
* [setDataSeriesNames()](#setdataseriesnames)
* [setDataSeriesTickLabels()](#setdataseriesticklabels) -- Backward compatible
* [getDisplayBlanksAs()](#getdisplayblanksas) -- Get Display Blanks as
* [setDisplayBlanksAs()](#setdisplayblanksas) -- Set Display Blanks as
* [getLegend()](#getlegend) -- Get Legend
* [setLegend()](#setlegend) -- Set Legend
* [setLegendPosition()](#setlegendposition)
* [setLegendPositionBottom()](#setlegendpositionbottom)
* [setLegendPositionLeft()](#setlegendpositionleft)
* [setLegendPositionRight()](#setlegendpositionright)
* [setLegendPositionTop()](#setlegendpositiontop)
* [getMajorGridlines()](#getmajorgridlines) -- Get Major Gridlines
* [getMinorGridlines()](#getminorgridlines) -- Get Minor Gridlines
* [getName()](#getname) -- Get Name
* [setName()](#setname)
* [getPlotArea()](#getplotarea) -- Get Plot Area
* [setPlotArea()](#setplotarea)
* [getPlotChartTypes()](#getplotcharttypes) -- Get the data series type(s) for a chart plot series
* [setPlotLayout()](#setplotlayout)
* [setPlotShowPercent()](#setplotshowpercent)
* [setPlotShowValues()](#setplotshowvalues)
* [getPlotVisibleOnly()](#getplotvisibleonly) -- Get Plot Visible Only
* [setPlotVisibleOnly()](#setplotvisibleonly) -- Set Plot Visible Only
* [setPosition()](#setposition) -- Set the Bottom Right position of the chart
* [setSheet()](#setsheet)
* [getTitle()](#gettitle) -- Get Title
* [setTitle()](#settitle) -- Set Title
* [getTopLeftCell()](#gettopleftcell) -- Get the cell address where the top left of the chart is fixed
* [setTopLeftCell()](#settopleftcell) -- Set the Top Left cell position for the chart
* [getTopLeftOffset()](#gettopleftoffset) -- Get the offset position within the Top Left cell for the chart
* [setTopLeftOffset()](#settopleftoffset) -- Set the offset position within the Top Left cell for the chart
* [getTopLeftPosition()](#gettopleftposition) -- Get the top left position of the chart
* [setTopLeftPosition()](#settopleftposition) -- Set the Top Left position for the chart
* [getTopLeftXOffset()](#gettopleftxoffset)
* [setTopLeftXOffset()](#settopleftxoffset)
* [getTopLeftYOffset()](#gettopleftyoffset)
* [setTopLeftYOffset()](#settopleftyoffset)
* [getValueAxisTitle()](#getvalueaxistitle) -- Get Value Axis Title
* [setValueAxisTitle()](#setvalueaxistitle) -- Set Value Axis Title
* [getValueAxisTitle2()](#getvalueaxistitle2) -- Get Y2 Axis Title
* [setValueAxisTitle2()](#setvalueaxistitle2) -- Set Y2 Axis Title

---

## __construct()

---

```php
public function __construct($title, $plotArea, ?Legend $legend = null, 
                            ?bool $plotVisibleOnly = true, 
                            ?string $displayBlanksAs = '0', $xAxisLabel, 
                            $yAxisLabel, ?Axis $xAxis = null, 
                            ?Axis $yAxis = null, 
                            ?GridLines $majorGridlines = null, 
                            ?GridLines $minorGridlines = null)
```
_Chart constructor_

### Parameters

* `Title|string $title`
* `PlotArea|array $plotArea`
* `Legend|null $legend`
* `bool|null $plotVisibleOnly`
* `string|null $displayBlanksAs`
* `Title|string|null $xAxisLabel`
* `Title|string|null $yAxisLabel`
* `Axis|null $xAxis`
* `Axis|null $yAxis`
* `GridLines|null $majorGridlines`
* `GridLines|null $minorGridlines`

---

## make()

---

```php
public static function make(string $chartType, $title, $dataSource): Chart
```


### Parameters

* `string $chartType`
* `Title|string $title`
* `DataSeries|array $dataSource`

---

## addDataSeriesSet()

---

```php
public function addDataSeriesSet(array $dataSources): Chart
```


### Parameters

* `array $dataSources`

---

## addDataSeriesType()

---

```php
public function addDataSeriesType(string $chartType, $dataSource, 
                                  ?string $dataLabel = null, 
                                  ?array $options = []): Chart
```


### Parameters

* `string $chartType`
* `DataSeriesValues|string $dataSource`
* `string|null $dataLabel`
* `array|null $options`

---

## addDataSeriesType2()

---

```php
public function addDataSeriesType2(string $chartType, $dataSource, 
                                   ?string $dataLabel = null, 
                                   ?array $options = []): Chart
```


### Parameters

* `string $chartType`
* `DataSeriesValues|string $dataSource`
* `string|null $dataLabel`
* `array|null $options`

---

## addDataSeriesValues()

---

```php
public function addDataSeriesValues($dataSource, ?string $dataLabel = null, 
                                    ?array $options = []): Chart
```


### Parameters

* `$dataSource`
* `string|null $dataLabel`
* `array|null $options`

---

## getBottomRightCell()

---

```php
public function getBottomRightCell(): string
```
_Get the cell address where the bottom right of the chart is fixed_

### Parameters

_None_

---

## setBottomRightCell()

---

```php
public function setBottomRightCell(string $cell): Chart
```


### Parameters

* `string $cell`

---

## getBottomRightOffset()

---

```php
public function getBottomRightOffset(): array
```
_Get the offset position within the Bottom Right cell for the chart_

### Parameters

_None_

---

## setBottomRightOffset()

---

```php
public function setBottomRightOffset(?int $xOffset = null, 
                                     ?int $yOffset = null): Chart
```
_Set the offset position within the Bottom Right cell for the chart_

### Parameters

* `int|null $xOffset`
* `int|null $yOffset`

---

## getBottomRightPosition()

---

```php
public function getBottomRightPosition(): array
```
_Get the bottom right position of the chart_

### Parameters

_None_

---

## getBottomRightXOffset()

---

```php
public function getBottomRightXOffset(): int
```


### Parameters

_None_

---

## setBottomRightXOffset()

---

```php
public function setBottomRightXOffset(int $xOffset): Chart
```


### Parameters

* `int $xOffset`

---

## getBottomRightYOffset()

---

```php
public function getBottomRightYOffset(): int
```


### Parameters

_None_

---

## setBottomRightYOffset()

---

```php
public function setBottomRightYOffset(int $yOffset): Chart
```


### Parameters

* `int $yOffset`

---

## setCategoryAxis()

---

```php
public function setCategoryAxis($labels, $title): Chart
```


### Parameters

* `$labels`
* `$title`

---

## setCategoryAxisLabels()

---

```php
public function setCategoryAxisLabels($labels): Chart
```
_Set Category Axis Labels_

### Parameters

* `$labels`

---

## getCategoryAxisTitle()

---

```php
public function getCategoryAxisTitle(): ?Title
```
_Get Category Axis Title_

### Parameters

_None_

---

## setCategoryAxisTitle()

---

```php
public function setCategoryAxisTitle($title): Chart
```
_Set Category Axis Title_

### Parameters

* `Title|string $title`

---

## getChartAxisX()

---

```php
public function getChartAxisX(): ?Axis
```
_Get xAxis_

### Parameters

_None_

---

## getChartAxisY()

---

```php
public function getChartAxisY(): ?Axis
```
_Get yAxis_

### Parameters

_None_

---

## getChartAxisY2()

---

```php
public function getChartAxisY2(): ?Axis
```
_Get yAxis_

### Parameters

_None_

---

## setChartColors()

---

```php
public function setChartColors(array $colors): Chart
```


### Parameters

* `array $colors`

---

## setChartType()

---

```php
public function setChartType(string $chartType): Chart
```


### Parameters

* `string $chartType`

---

## setDataSeriesNames()

---

```php
public function setDataSeriesNames($labels): Chart
```


### Parameters

* `$labels`

---

## setDataSeriesTickLabels()

---

```php
public function setDataSeriesTickLabels($range): Chart
```
_Backward compatible_

### Parameters

* `$range`

---

## getDisplayBlanksAs()

---

```php
public function getDisplayBlanksAs(): string
```
_Get Display Blanks as_

### Parameters

_None_

---

## setDisplayBlanksAs()

---

```php
public function setDisplayBlanksAs(string $displayBlanksAs = '0')
```
_Set Display Blanks as_

### Parameters

* `string $displayBlanksAs`

---

## getLegend()

---

```php
public function getLegend(): ?Legend
```
_Get Legend_

### Parameters

_None_

---

## setLegend()

---

```php
public function setLegend(avadim\FastExcelWriter\Charts\Legend $legend): Chart
```
_Set Legend_

### Parameters

* `Legend $legend`

---

## setLegendPosition()

---

```php
public function setLegendPosition(string $position): Chart
```


### Parameters

* `string $position`

---

## setLegendPositionBottom()

---

```php
public function setLegendPositionBottom(): Chart
```


### Parameters

_None_

---

## setLegendPositionLeft()

---

```php
public function setLegendPositionLeft(): Chart
```


### Parameters

_None_

---

## setLegendPositionRight()

---

```php
public function setLegendPositionRight(): Chart
```


### Parameters

_None_

---

## setLegendPositionTop()

---

```php
public function setLegendPositionTop(): Chart
```


### Parameters

_None_

---

## getMajorGridlines()

---

```php
public function getMajorGridlines(): ?GridLines
```
_Get Major Gridlines_

### Parameters

_None_

---

## getMinorGridlines()

---

```php
public function getMinorGridlines(): ?GridLines
```
_Get Minor Gridlines_

### Parameters

_None_

---

## getName()

---

```php
public function getName(): string
```
_Get Name_

### Parameters

_None_

---

## setName()

---

```php
public function setName(string $name): Chart
```


### Parameters

* `string $name`

---

## getPlotArea()

---

```php
public function getPlotArea(): PlotArea
```
_Get Plot Area_

### Parameters

_None_

---

## setPlotArea()

---

```php
public function setPlotArea($plotValues): Chart
```


### Parameters

* `$plotValues`

---

## getPlotChartTypes()

---

```php
public function getPlotChartTypes(): array
```
_Get the data series type(s) for a chart plot series_

### Parameters

_None_

---

## setPlotLayout()

---

```php
public function setPlotLayout($layout): Chart
```


### Parameters

* `$layout`

---

## setPlotShowPercent()

---

```php
public function setPlotShowPercent(bool $val): Chart
```


### Parameters

* `bool $val`

---

## setPlotShowValues()

---

```php
public function setPlotShowValues(bool $val): Chart
```


### Parameters

* `bool $val`

---

## getPlotVisibleOnly()

---

```php
public function getPlotVisibleOnly(): bool
```
_Get Plot Visible Only_

### Parameters

_None_

---

## setPlotVisibleOnly()

---

```php
public function setPlotVisibleOnly(bool $plotVisibleOnly = true): Chart
```
_Set Plot Visible Only_

### Parameters

* `boolean $plotVisibleOnly`

---

## setPosition()

---

```php
public function setPosition(string $cell, ?int $xOffset = null, 
                            ?int $yOffset = null): Chart
```
_Set the Bottom Right position of the chart_

### Parameters

* `string $cell`
* `int|null $xOffset`
* `int|null $yOffset`

---

## setSheet()

---

```php
public function setSheet(avadim\FastExcelWriter\Sheet $sheet): Chart
```


### Parameters

* `Sheet $sheet`

---

## getTitle()

---

```php
public function getTitle(): Title
```
_Get Title_

### Parameters

_None_

---

## setTitle()

---

```php
public function setTitle($title): Chart
```
_Set Title_

### Parameters

* `Title|string $title`

---

## getTopLeftCell()

---

```php
public function getTopLeftCell(): string
```
_Get the cell address where the top left of the chart is fixed_

### Parameters

_None_

---

## setTopLeftCell()

---

```php
public function setTopLeftCell(string $cell): Chart
```
_Set the Top Left cell position for the chart_

### Parameters

* `string $cell`

---

## getTopLeftOffset()

---

```php
public function getTopLeftOffset(): array
```
_Get the offset position within the Top Left cell for the chart_

### Parameters

_None_

---

## setTopLeftOffset()

---

```php
public function setTopLeftOffset(?int $xOffset = null, 
                                 ?int $yOffset = null): Chart
```
_Set the offset position within the Top Left cell for the chart_

### Parameters

* `integer|null $xOffset`
* `integer|null $yOffset`

---

## getTopLeftPosition()

---

```php
public function getTopLeftPosition(): array
```
_Get the top left position of the chart_

### Parameters

_None_

---

## setTopLeftPosition()

---

```php
public function setTopLeftPosition(string $cell, ?int $xOffset = null, 
                                   ?int $yOffset = null): Chart
```
_Set the Top Left position for the chart_

### Parameters

* `string $cell`
* `integer|null $xOffset`
* `integer|null $yOffset`

---

## getTopLeftXOffset()

---

```php
public function getTopLeftXOffset(): int
```


### Parameters

_None_

---

## setTopLeftXOffset()

---

```php
public function setTopLeftXOffset($xOffset): Chart
```


### Parameters

* `$xOffset`

---

## getTopLeftYOffset()

---

```php
public function getTopLeftYOffset(): int
```


### Parameters

_None_

---

## setTopLeftYOffset()

---

```php
public function setTopLeftYOffset($yOffset): Chart
```


### Parameters

* `$yOffset`

---

## getValueAxisTitle()

---

```php
public function getValueAxisTitle(): ?Title
```
_Get Value Axis Title_

### Parameters

_None_

---

## setValueAxisTitle()

---

```php
public function setValueAxisTitle($title): Chart
```
_Set Value Axis Title_

### Parameters

* `Title|string $title`

---

## getValueAxisTitle2()

---

```php
public function getValueAxisTitle2(): ?Title
```
_Get Y2 Axis Title_

### Parameters

_None_

---

## setValueAxisTitle2()

---

```php
public function setValueAxisTitle2($title): Chart
```
_Set Y2 Axis Title_

### Parameters

* `Title|string $title`

---

