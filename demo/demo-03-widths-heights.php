<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$timer = microtime(true);
$excel = Excel::create();
$sheet = $excel->getSheet();

$sheet->setColWidths([10, 20, 30, 40]);
$sheet->setRowHeight(2, 33);
//$sheet->setRowHeights([1 => 20, 2 => 33, 3 => 40]);

$sheet->writeRow([300, 234, 456, 789], ['height' => 20]);
$sheet->writeRow([300, 234, 456, 789]);
$sheet->writeRow([300, 234, 456, 789], ['height' => 40]);

$excel->save($outFileName);

echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF