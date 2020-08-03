<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$chars = 'abcdefghijklmnopqrstuvwxyz0123456789 ';
$s = '';
for ($j = 0; $j < 16192; $j++) {
    $s .= $chars[mt_rand() % 36];
}

$s1 = substr($s, mt_rand() % 4000, mt_rand() % 5 + 5);
$s2 = substr($s, mt_rand() % 8000, mt_rand() % 5 + 5);
$s3 = substr($s, mt_rand() % 12000, mt_rand() % 5 + 5);
$s4 = substr($s, mt_rand() % 16000, mt_rand() % 5 + 5);
$row = [$s1, $s2, $s3, $s4];


$timer = microtime(true);
$excel = Excel::create(['250K']);
$sheet = $excel->getSheet();

$sheet->setColFormats(['string', 'string', 'string', 'string']);

$rowCount = 250000;
for($i = 0; $i < $rowCount; $i++) {
    $sheet->writeRow($row);
}

$excel->save($outFileName);

echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF