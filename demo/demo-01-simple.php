<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$header = [
    'c1-text'   => 'text',//text
    'c2-text'   => '@',//text
    'c3-integer' => 'integer',
    //'c4-integer' => '0',
    //'c5-money'  => 'money',
    'c4-integer' => ['text-color' => '#f00'],
    'c5-money'  => '#,##0.00 [$₽];[RED]-#,##0.00 [$₽]', //'money', //['format' => 'money', 'font' => 'bold'],
    'c6-price'  => '#,##0.00',//custom
    //'c7-date'   => 'date',
    'c7-date'   => ['format' => 'date', 'width' => 11],
    'c8-date'   => ['format' => 'YYYY-MM-DD', 'width' => 11],
    'c9-time'   => 'H:MM',
];

$data = [
    ['Moscow',      102, 103, 104, 1, 1106, '2018-01-08', '2018-01-08', '1:04'],
    ['Paris',       202, 203, 204, 12, 2206, '2018-02-08', '2018-02-08', '1:59'],
    ['Washington',  302, 303, 304, 123, 3306, '2018-03-08', '2018-03-08', '13:59'],
    ['Berlin',      402, 403, 404, 1234, 4406, '2018-04-08', '2018-04-08', '1:59 am'],
    ['Rome',        502, 503, 504, -123, 5506, '2018-05-08', '2018-05-08', '1:59 pm'],
    ['Santiago',    602, 603, 604, -12, 6606, '2018-06-08', '2018-06-08', '23:17'],
    ['Canberra',    702, 703, 704, 0, 7706, '2018-07-08', '2018-07-08', '7:10'],
];

$timer = microtime(true);
$excel = Excel::create();
$sheet = $excel->getSheet();

// The fastest way to write data is row by row
$sheet->writeHeader($header);

foreach($data as $row) {
    $sheet->writeRow($row);
}

$excel->save($outFileName);

echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF