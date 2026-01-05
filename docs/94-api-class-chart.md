# Class \avadim\FastExcelWriter\Charts\Chart

---

* [__construct()](#__construct) -- Chart constructor
* [make()](#make) -- Create Chart instance
* [addDataSeriesSet()](#adddataseriesset) -- Add multiple data series to the chart
* [addDataSeriesType()](#adddataseriestype) -- Add data series of specific type (used for combo charts)
* [addDataSeriesType2()](#adddataseriestype2) -- Add data series of specific type to the second axis (used for combo charts)
* [addDataSeriesValues()](#adddataseriesvalues) -- Add data series values to the chart
* [getBottomRightCell()](#getbottomrightcell) -- Get the cell address where the bottom right of the chart is fixed
* [setBottomRightCell()](#setbottomrightcell) -- Set Bottom Right Cell
* [getBottomRightOffset()](#getbottomrightoffset) -- Get the offset position within the Bottom Right cell for the chart
* [setBottomRightOffset()](#setbottomrightoffset) -- Set the offset position within the Bottom Right cell for the chart
* [getBottomRightPosition()](#getbottomrightposition) -- Get the bottom right position of the chart
* [getBottomRightXOffset()](#getbottomrightxoffset) -- Get Bottom Right X-Offset
* [setBottomRightXOffset()](#setbottomrightxoffset) -- Set Bottom Right X-Offset
* [getBottomRightYOffset()](#getbottomrightyoffset) -- Get Bottom Right Y-Offset
* [setBottomRightYOffset()](#setbottomrightyoffset) -- Set Bottom Right Y-Offset
* [setCategoryAxis()](#setcategoryaxis) -- Set Category Axis Labels and Title
* [setCategoryAxisLabels()](#setcategoryaxislabels) -- Set Category Axis Labels (tick labels)
* [getCategoryAxisTitle()](#getcategoryaxistitle) -- Get Category Axis Title
* [setCategoryAxisTitle()](#setcategoryaxistitle) -- Set Category Axis Title
* [getChartAxisX()](#getchartaxisx) -- Get X Axis
* [getChartAxisY()](#getchartaxisy) -- Get Y Axis
* [getChartAxisY2()](#getchartaxisy2) -- Get Y Axis 2
* [setChartColors()](#setchartcolors) -- Set Chart Colors
* [setChartType()](#setcharttype) -- Set Chart Type
* [setDataSeriesNames()](#setdataseriesnames) -- Set Data Series Names
* [setDataSeriesTickLabels()](#setdataseriesticklabels) -- Set Data Series Tick Labels
* [getDisplayBlanksAs()](#getdisplayblanksas) -- Get Display Blanks As
* [setDisplayBlanksAs()](#setdisplayblanksas) -- Set Display Blanks As
* [getLegend()](#getlegend) -- Get Chart Legend
* [setLegend()](#setlegend) -- Set Chart Legend
* [setLegendPosition()](#setlegendposition) -- Set Chart Legend Position
* [setLegendPositionBottom()](#setlegendpositionbottom) -- Set Chart Legend Position to Bottom
* [setLegendPositionLeft()](#setlegendpositionleft) -- Set Chart Legend Position to Left
* [setLegendPositionRight()](#setlegendpositionright) -- Set Chart Legend Position to Right
* [setLegendPositionTop()](#setlegendpositiontop) -- Set Chart Legend Position to Top
* [getMajorGridlines()](#getmajorgridlines) -- Get Major Gridlines
* [getMinorGridlines()](#getminorgridlines) -- Get Minor Gridlines
* [getName()](#getname) -- Get Chart Name
* [setName()](#setname) -- Set Chart Name
* [getPlotArea()](#getplotarea) -- Get Plot Area
* [setPlotArea()](#setplotarea) -- Set Chart Plot Area
* [getPlotChartTypes()](#getplotcharttypes) -- Get the data series type(s) for a chart plot series
* [setPlotLayout()](#setplotlayout) -- Set Plot Layout
* [setPlotShowPercent()](#setplotshowpercent) -- Set Plot Show Percent
* [setPlotShowValues()](#setplotshowvalues) -- Set Plot Show Values
* [getPlotVisibleOnly()](#getplotvisibleonly) -- Get Plot Visible Only
* [setPlotVisibleOnly()](#setplotvisibleonly) -- Set Plot Visible Only
* [setPosition()](#setposition) -- Set the Bottom Right position of the chart
* [setSheet()](#setsheet) -- Set Sheet
* [getTitle()](#gettitle) -- Get Chart Title
* [setTitle()](#settitle) -- Set Chart Title
* [getTopLeftCell()](#gettopleftcell) -- Get the cell address where the top left of the chart is fixed
* [setTopLeftCell()](#settopleftcell) -- Set the Top Left cell position for the chart
* [getTopLeftOffset()](#gettopleftoffset) -- Get the offset position within the Top Left cell for the chart
* [setTopLeftOffset()](#settopleftoffset) -- Set the offset position within the Top Left cell for the chart
* [getTopLeftPosition()](#gettopleftposition) -- Get the top left position of the chart
* [setTopLeftPosition()](#settopleftposition) -- Set the Top Left position for the chart
* [getTopLeftXOffset()](#gettopleftxoffset) -- Get Top Left X-Offset
* [setTopLeftXOffset()](#settopleftxoffset) -- Set Top Left X-Offset
* [getTopLeftYOffset()](#gettopleftyoffset) -- Get Top Left Y-Offset
* [setTopLeftYOffset()](#settopleftyoffset) -- Set Top Left Y-Offset
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
_Create Chart instance_

### Parameters

* `string $chartType`
* `Title|string|null $title`
* `DataSeries|PlotArea|array|null $dataSource`

---

## addDataSeriesSet()

---

```php
public function addDataSeriesSet(array $dataSources): Chart
```
_Add multiple data series to the chart_

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
_Add data series of specific type (used for combo charts)_

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
_Add data series of specific type to the second axis (used for combo charts)_

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
_Add data series values to the chart_

### Parameters

* `DataSeriesValues|array|string $dataSource`
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
_Set Bottom Right Cell_

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
_Get Bottom Right X-Offset_

### Parameters

_None_

---

## setBottomRightXOffset()

---

```php
public function setBottomRightXOffset(int $xOffset): Chart
```
_Set Bottom Right X-Offset_

### Parameters

* `int $xOffset`

---

## getBottomRightYOffset()

---

```php
public function getBottomRightYOffset(): int
```
_Get Bottom Right Y-Offset_

### Parameters

_None_

---

## setBottomRightYOffset()

---

```php
public function setBottomRightYOffset(int $yOffset): Chart
```
_Set Bottom Right Y-Offset_

### Parameters

* `int $yOffset`

---

## setCategoryAxis()

---

```php
public function setCategoryAxis($labels, $title): Chart
```
_Set Category Axis Labels and Title_

### Parameters

* `DataSeriesValues|array|string $labels`
* `Title|string|null $title`

---

## setCategoryAxisLabels()

---

```php
public function setCategoryAxisLabels($labels): Chart
```
_Set Category Axis Labels (tick labels)_

### Parameters

* `DataSeriesValues|array|string $labels`

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
_Get X Axis_

### Parameters

_None_

---

## getChartAxisY()

---

```php
public function getChartAxisY(): ?Axis
```
_Get Y Axis_

### Parameters

_None_

---

## getChartAxisY2()

---

```php
public function getChartAxisY2(): ?Axis
```
_Get Y Axis 2_

### Parameters

_None_

---

## setChartColors()

---

```php
public function setChartColors(array $colors): Chart
```
_Set Chart Colors_

### Parameters

* `array $colors`

---

## setChartType()

---

```php
public function setChartType(string $chartType): Chart
```
_Set Chart Type_

### Parameters

* `string $chartType`

---

## setDataSeriesNames()

---

```php
public function setDataSeriesNames($labels): Chart
```
_Set Data Series Names_

### Parameters

* `DataSeriesValues|array|string $labels`

---

## setDataSeriesTickLabels()

---

```php
public function setDataSeriesTickLabels($range): Chart
```
_Set Data Series Tick Labels_

### Parameters

* `DataSeriesValues|array|string $range`

---

## getDisplayBlanksAs()

---

```php
public function getDisplayBlanksAs(): string
```
_Get Display Blanks As_

### Parameters

_None_

---

## setDisplayBlanksAs()

---

```php
public function setDisplayBlanksAs(string $displayBlanksAs = '0'): Chart
```
_Set Display Blanks As_

### Parameters

* `string $displayBlanksAs`

---

## getLegend()

---

```php
public function getLegend(): ?Legend
```
_Get Chart Legend_

### Parameters

_None_

---

## setLegend()

---

```php
public function setLegend(avadim\FastExcelWriter\Charts\Legend $legend): Chart
```
_Set Chart Legend_

### Parameters

* `Legend $legend`

---

## setLegendPosition()

---

```php
public function setLegendPosition(string $position): Chart
```
_Set Chart Legend Position_

### Parameters

* `string $position`

---

## setLegendPositionBottom()

---

```php
public function setLegendPositionBottom(): Chart
```
_Set Chart Legend Position to Bottom_

### Parameters

_None_

---

## setLegendPositionLeft()

---

```php
public function setLegendPositionLeft(): Chart
```
_Set Chart Legend Position to Left_

### Parameters

_None_

---

## setLegendPositionRight()

---

```php
public function setLegendPositionRight(): Chart
```
_Set Chart Legend Position to Right_

### Parameters

_None_

---

## setLegendPositionTop()

---

```php
public function setLegendPositionTop(): Chart
```
_Set Chart Legend Position to Top_

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
_Get Chart Name_

### Parameters

_None_

---

## setName()

---

```php
public function setName(string $name): Chart
```
_Set Chart Name_

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
_Set Chart Plot Area_

### Parameters

* `PlotArea|array $plotValues`

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
_Set Plot Layout_

### Parameters

* `Layout|array|null $layout`

---

## setPlotShowPercent()

---

```php
public function setPlotShowPercent(bool $val): Chart
```
_Set Plot Show Percent_

### Parameters

* `bool $val`

---

## setPlotShowValues()

---

```php
public function setPlotShowValues(bool $val): Chart
```
_Set Plot Show Values_

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

* `bool $plotVisibleOnly`

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
_Set Sheet_

### Parameters

* `Sheet $sheet`

---

## getTitle()

---

```php
public function getTitle(): Title
```
_Get Chart Title_

### Parameters

_None_

---

## setTitle()

---

```php
public function setTitle($title): Chart
```
_Set Chart Title_

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
_Get Top Left X-Offset_

### Parameters

_None_

---

## setTopLeftXOffset()

---

```php
public function setTopLeftXOffset($xOffset): Chart
```
_Set Top Left X-Offset_

### Parameters

* `int $xOffset`

---

## getTopLeftYOffset()

---

```php
public function getTopLeftYOffset(): int
```
_Get Top Left Y-Offset_

### Parameters

_None_

---

## setTopLeftYOffset()

---

```php
public function setTopLeftYOffset($yOffset): Chart
```
_Set Top Left Y-Offset_

### Parameters

* `int $yOffset`

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

