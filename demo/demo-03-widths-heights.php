<?php

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$timer = microtime(true);
$excel = Excel::create();
$sheet = $excel->getSheet();

$heights = [
    2 => 15,
    3 => 30,
    4 => 60,
];

$widths = ['A' => null, 'B' => 40, 'C' => 30, 'D' => 20, 'E' => 10];

$values = [];
foreach ($widths as $w) {
    $values[] = $w ? ('width:' . $w) : null;
}
$sheet->writeRow($values, ['text-align' => 'center', 'font' => 'bold', 'border' => 'thin']);
$sheet->setColWidths($widths);

// set style foe the cell A only
$cellStyles = [['font' => 'bold', 'border' => 'thin']];
$sheet->setRowHeight(2, $heights[2]);
$sheet->writeRow(['height:' . $heights[2] ?? '', 234, 456, 789], [], $cellStyles);

// Write row with specified height
$sheet->writeRow(['height:' . $heights[3] ?? '', 234, 456, 789], ['height' => $heights[3]], $cellStyles);

// Write row with specified height - other way (preferred)
$sheet->writeRow(['height:' . $heights[4] ?? '', 234, 456, 789], [], $cellStyles)
    ->applyRowHeight($heights[4]);

$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF