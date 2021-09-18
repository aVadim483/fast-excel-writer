<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$chars = 'abcdefghijklmnopqrstuvwxyz0123456789 ';
$data = [];
for($i=0; $i<1000; $i++) {
    $data[] = [
        $i,
        str_shuffle($chars),
        mt_rand() % 10000,
        date('Y-m-d H:i:s',time() - (mt_rand() % 31536000)),
    ];
}

$timer = microtime(true);
$excel = Excel::create();
$sheet = $excel->getSheet();

$sheet->setColWidths([8, 36, 8, 18]);
$sheet->writeHeader(['Num' => 'integer', 'Str' => 'string', 'Float' => '0.00', 'Date' => 'datetime']);
foreach($data as $row) {
    $sheet->writeRow($row);
}

$sheet
    ->setFreeze('B2')
    ->setAutofilter(1);

$excel->save($outFileName);

echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF