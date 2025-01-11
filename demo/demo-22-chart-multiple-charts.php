<?php

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use avadim\FastExcelWriter\Charts\Chart;
use avadim\FastExcelWriter\Charts\Legend;
use \avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\StyleManager;

$timer = microtime(true);

// Create Excel workbook
$excel = Excel::create(['Chart Demo']);

$sheet = $excel->sheet();

$data = [
    ['',	        2010,	    2011,	        2012],
    ['Q1',          12,         15,		        21],
    ['Q2',          56,         73,		        86],
    ['Q3',          52,         61,		        69],
    ['Q4',          30,         32,		        0],
    ['Total', '=SUM(B2:B5)', '=SUM(C2:C5)', '=SUM(D2:D5)'],
];

foreach ($data as $row) {
    $sheet->writeRow($row);
}

$chart = Chart::make(Chart::TYPE_COLUMN, 'Column Chart', ['b1' => 'B2:B5', 'c1' => 'c2:c5', 'd1' => 'd2:d5'])
    ->setCategoryAxisLabels('A2:A5')
    ->setLegendPosition(Legend::POSITION_TOPRIGHT)
;

//	Add the chart to the worksheet
$sheet->addChart('a9:h22', $chart);

$chart = Chart::make(Chart::TYPE_PIE, 'Pie Chart', ['b6:d6'])
    ->setCategoryAxisLabels('b1:d1')
    ->setLegendPosition(Legend::POSITION_TOPRIGHT)
    ->setPlotShowPercent(true)
;
$sheet->addChart('i9:m22', $chart);

$chart = Chart::make(Chart::TYPE_BAR_STACKED, 'Stacked Bar Chart', ['b1' => 'B2:B5', 'c1' => 'c2:c5', 'd1' => 'd2:d5'])
    ->setCategoryAxisLabels('A2:A5')
    ->setLegendPosition(Legend::POSITION_TOPRIGHT)
;
$sheet->addChart('a23:h36', $chart);

$chart = Chart::make(Chart::TYPE_PIE_3D, '3D Pie Chart', ['b6:d6'])
    ->setCategoryAxisLabels('b1:d1')
    ->setLegendPosition(Legend::POSITION_TOPRIGHT)
    ->setPlotShowPercent(true)
;
$sheet->addChart('i23:m36', $chart);

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF