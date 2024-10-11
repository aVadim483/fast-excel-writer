<?php

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$data = [
    '1985-01-28 23:05:59',
    '1985-01-28',
    '23:05:59',
    '23:05',
];

$formats = [
    '@',
    '@datetime',
    '@date',
    '@time',
    'YYYY-MM-DD HH:MM:SS',
    'DD MMM YY',
    'H:MM'
];

$timer = microtime(true);
$excel = Excel::create();
$sheet = $excel->getSheet();

$columns = [];
$colNum = 1;
foreach ($formats as $format) {
    $columns[$colNum++] = [
        'format' => $format,
        'width' => 24,
    ];
}
$sheet->setColDataStyleArray($columns);
$sheet->writeHeader($formats, ['font' => 'bold', 'text-align' => 'center', 'border' => 'thin']);

foreach ($data as $value) {
    foreach ($formats as $format) {
        // write values in one row cell by cell
        $sheet->writeCell($value);
    }
    // go to the first cell of the next row
    $sheet->nextRow();
}

$currentTime = time();
foreach ($formats as $format) {
    $sheet->writeCell($currentTime, ['fill' => '#eee']);
}

$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF