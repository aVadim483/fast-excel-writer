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
    ['',		'',		'Budget',	'Forecast',	'Actual'],
    ['2010',	'Q1',   47,   		44,			43],
    ['',		'Q2',   56,   		53,			50],
    ['',		'Q3',   52,   		46,			45],
    ['',		'Q4',   45,   		40,			40],
    ['2011',	'Q1',   51,   		42,			46],
    ['',		'Q2',   53,   		58,			56],
    ['',		'Q3',   64,   		66,			69],
    ['',		'Q4',   54,   		55,			56],
    ['2012',	'Q1',   49,   		52,			58],
    ['',		'Q2',   68,   		73,			86],
    ['',		'Q3',   72,   		78,			0],
    ['',		'Q4',   50,   		60,			0],
];

foreach ($data as $row) {
    $sheet->writeRow($row);
}

$chart1 = Chart::make(Chart::TYPE_COLUMN, 'Grouped Column Chart', ['c1' => 'c2:c13', 'd1' => 'd2:d13', 'e1' => 'e2:e13'])
    ->setDataSeriesTickLabels('A2:B13')
    ->setLegendPosition(Legend::POSITION_BOTTOM)
    ->setYAxisLabel('Value ($k)')
    ->setXAxisLabel('Financial Period')
;

//	Add the chart to the worksheet
$sheet->addChart('G2:P20', $chart1);

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF