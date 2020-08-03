<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$timer = microtime(true);
$excel = Excel::create(['Total']);

// make new sheet
$sheet = $excel->makeSheet('Jan');
$sheet->setColFormat(1, 'date');
for ($day = 1; $day <= 31; $day++) {
    $sheet->writeRow(['2020-1-' . $day, random_int(100, 999)]);
}

// make new sheet
$sheet = $excel->makeSheet('Feb');
$sheet->setColFormat(1, 'date');
for ($day = 1; $day <= 29; $day++) {
    $sheet->writeRow(['2020-2-' . $day, random_int(100, 999)]);
}

// make new sheet
$sheet = $excel->makeSheet('Mar');
$sheet->setColFormat(1, 'date');
for ($day = 1; $day <= 31; $day++) {
    $sheet->writeRow(['2020-3-' . $day, random_int(100, 999)]);
}

$sheet = $excel->getSheet('Total');

$sheet->writeRow(['Jan 2020', '=SUM(Jan!B:B)']);
$sheet->writeRow(['Feb 2020', '=SUM(Feb!B:B)']);
$sheet->writeRow(['Mar 2020', '=SUM(Mar!B:B)']);

$excel->save($outFileName);

echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF