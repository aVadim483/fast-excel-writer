<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$timer = microtime(true);
$excel = Excel::create();
$sheet = $excel->getSheet();

$heights = [
    1 => 12,
    2 => 25,
    3 => 50,
];

//$sheet->setColWidths([10, 20, 30, 40]);
$sheet->setColWidths(['A' => 40, 'B' => 30, 'C' => 20, 'D' => 10]);
$sheet->setRowHeight(2, $heights[2]);
//$sheet->setRowHeights([1 => 20, 2 => 33, 3 => 40]);

$rowNum = 0;
// Write row width default height
$sheet->writeRow(['height: ' . $heights[++$rowNum] ?? '', 234, 456, 789]);

// Write row with specified height
$sheet->writeRow(['height: ' . $heights[++$rowNum] ?? '', 234, 456, 789]);

// Write row with specified height - other way (preferred)
$sheet->writeRow(['height: ' . $heights[++$rowNum] ?? '', 234, 456, 789])
    ->applyRowHeight($heights[$rowNum]);

$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF