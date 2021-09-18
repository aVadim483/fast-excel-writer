<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$header = [
    'c1-text'   => 'text',//text
    'c2-text'   => '@',//text
    'c3-integer' => 'integer',
    'c4-integer' => '0',
    'c5-money'  => 'money',
    'c6-price'  => '#,##0.00',//custom
    'c7-date'   => 'date',
    'c8-date'   => 'YYYY-MM-DD',
];

$data = [
    ['Moscow',      102, 103, 104, 105, 106, '2018-01-07', '2018-01-08'],
    ['Paris',       202, 203, 204, 205, 206, '2018-02-07', '2018-02-08'],
    ['Washington',  302, 303, 304, 305, 306, '2018-03-07', '2018-03-08'],
    ['Berlin',      402, 403, 404, 405, 406, '2018-04-07', '2018-04-08'],
    ['Rome',        502, 503, 504, 505, 506, '2018-05-07', '2018-05-08'],
    ['Santiago',    602, 603, 604, 605, 606, '2018-06-07', '2018-06-08'],
    ['Canberra',    702, 703, 704, 705, 706, '2018-07-07', '2018-07-08'],
];

$timer = microtime(true);
$excel = Excel::create();
$excel->setRightToLeft(true);

$sheet = $excel->getSheet();

// The fastest way to write data is row by row
$sheet->writeHeader($header);

foreach($data as $row) {
    $sheet->writeRow($row);
}

$excel->save($outFileName);

echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF