<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$header = [
    'c1-text'   => '@', //text
    'c2-text'   => '@text', //text
    'c3-integer' => '@integer',
    'c4-integer' => ['text-color' => '#f0f'], // default format
    'c5-money'  => '@money',
    'c6-price'  => '#\'##0.000', //custom numeric format
    'c7-date'   => ['format' => '@date', 'width' => 'auto'],
    'c8-date'   => ['format' => 'YYYY-MM-DD', 'width' => 11],
    'c9-time'   => 'H:MM',
];

$data = [
    // column A    B    C    D    E     F     G             H             I
    ['Moscow',     102, 103, 104, 1,    1106, '2018-01-08', '2018-01-08', '1:04'],    // row 2
    ['Paris',      202, 203, 204, 12,   2206, '2018-02-08', '2018-02-08', '1:59'],    // row 3
    ['Washington', 302, 303, 304, 123,  3306, '2018-03-08', '2018-03-08', '13:59'],   // row 4
    ['Berlin',     402, 403, 404, 1234, 4406, '2018-04-08', '2018-04-08', '1:59 am'], // row 5
    ['Rome',       502, 503, 504, -123, 5506, '2018-05-08', '2018-05-08', '1:59 pm'], // row 6
    ['Santiago',   602, 603, 604, -12,  6606, '2018-06-08', '2018-06-08', '23:17'],   // row 7
    ['Canberra',   702, 703, 704, 0,    7706, '2018-07-08', '2018-07-08', '7:10'],    // row 8
];

$timer = microtime(true);

// Create Excel workbook
$excel = Excel::create();

// Get the first sheet;
$sheet = $excel->getSheet();

$rowOptions = ['font-style' => 'bold'];
// Write header:
//    $header - values for cells of the first row and options for columns
//    $rowOptions - options of header row
$sheet->writeHeader($header, $rowOptions);

// The fastest way to write data is row by row
foreach($data as $row) {
    $sheet->writeRow($row);
}

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF