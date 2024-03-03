## FastExcelWriter - Charts

You can insert charts to generated spreadsheets (You can find usage examples below or in */demo* folder)

### Simple usage of chart

```php
// Create Excel workbook
$excel = Excel::create(['Chart Demo']);

$sheet = $excel->sheet();

$data = [
    ['',	2010,	2011,	2012],
    ['Q1',   12,   15,		21],
    ['Q2',   56,   73,		86],
    ['Q3',   52,   61,		69],
    ['Q4',   30,   32,		0],
];

foreach ($data as $row) {
    $sheet->writeRow($row);
}

// Define data series for chart
$dataSeries = [
    // key - cell with name of data series
    // value - range with data series
    'B1' => 'B2:B5', 
    'C1' => 'c2:c5', 
    'D1' => 'd2:d5',
];

$chartTitle = 'Bar Chart';

// Create chart
$chart = Chart::make(Chart::TYPE_COLUMN, $chartTitle, $dataSeries)
    // X axis tick values
    ->setDataSeriesTickLabels('A2:A5')
    // Position of legend
    ->setLegendPosition(Legend::POSITION_TOPRIGHT)
;

//	Add the chart to the worksheet
$sheet->addChart('A7:H20', $chart);

// Save to XLSX-file
$excel->save($outFileName);

```

### Multiple charts

```php
// Make chart 1
$chart1 = Chart::make(Chart::TYPE_COLUMN, 'Column Chart', ['b1' => 'B2:B5', 'c1' => 'c2:c5', 'd1' => 'd2:d5'])
    ->setDataSeriesTickLabels('A2:A5')
    ->setLegendPosition(Legend::POSITION_TOPRIGHT)
;

//	Add the chart to the worksheet
$sheet->addChart('a9:h22', $chart1);

// Make chart 2
$chart2 = Chart::make(Chart::TYPE_PIE, 'Pie Chart', ['b6:d6'])
    ->setDataSeriesTickLabels('b1:d1')
    ->setLegendPosition(Legend::POSITION_TOPRIGHT)
    ->setPlotShowPercent(true)
;

//	Add the chart to the worksheet
$sheet->addChart('i9:m22', $chart2);
```

### Chart types

| const name of class<br/>Chart       | chart type            |                                                                         |
|-------------------------------------|-----------------------|-------------------------------------------------------------------------|
| TYPE_BAR                            | bar chart             | ![img/chart-bar_240.jpg](img/chart-bar_240.jpg)                         |
| TYPE_BAR_STACKED                    | stacked bar chart     | ![img/chart-bar_240.jpg](img/chart-bar_240.jpg)                         |
| TYPE_COLUMN                         | column chart          | ![img/chart-column_240.jpg](img/chart-column_240.jpg)                   |
| TYPE_COLUMN_STACKED                 | stacked column chart  | ![img/chart-column-stacked_240.jpg](img/chart-column-stacked_240.jpg)   |
| TYPE_LINE                           | line chart            | ![img/chart-line_240.jpg](img/chart-line_240.jpg)                       |
| TYPE_LINE_STACKED                   | stacked line chart    | ![img/chart-line-stacked_240.jpg](img/chart-line-stacked_240.jpg)       |
| TYPE_LINE_3D                        | line 3D chart         | ![img/chart-line-3d_240.jpg](img/chart-line-3d_240.jpg)                 |
| TYPE_LINE_3D_STACKED                | stacked line 3D chart | ![img/chart-line-3d-stacked_240.jpg](img/chart-line-3d-stacked_240.jpg) |
| TYPE_AREA                           | area chart            | ![img/chart-area_240.jpg](img/chart-area_240.jpg)                       |
| TYPE_AREA_STACKED                   | stacked area chart    | ![img/chart-area-stacked_240.jpg](img/chart-area-stacked_240.jpg)       |
| TYPE_AREA_3D                        | area 3D chart         | ![img/chart-area-3d_240.jpg](img/chart-area-3d_240.jpg)                 |
| TYPE_AREA_3D_STACKED                | stacked area 3D chart | ![img/chart-area-stacked_240.jpg](img/chart-area-stacked_240.jpg)       |
| TYPE_PIE                            | pie chart             | ![img/chart-pie_240.jpg](img/chart-pie_240.jpg)                         |
| TYPE_PIE_3D                         | pie 3D chart          | ![img/chart-pie-3d_240.jpg](img/chart-pie-3d_240.jpg)                   |
| TYPE_DONUT                          | doughnut chart        | ![img/chart-donut_240.jpg](img/chart-donut_240.jpg)                     |

### Useful Chart Methods 

* setDataSeriesTickLabels(\<range>) - X axis tick values
* setLegendPosition(\<position>) - position of legend (use constants Legend::POSITION_XXX)
* setPlotShowValues(true) - show values on the chart
* setPlotShowPercent(true) - show values is percents (for pie and sonut)
