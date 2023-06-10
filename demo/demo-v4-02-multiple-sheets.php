<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$timer = microtime(true);
$excel = Excel::create(['Total']);

// make new sheet with name 'Jan'
$sheet1 = $excel->makeSheet('Jan');
$sheet1->setColFormat(1, '@date');
for ($day = 1; $day <= 31; $day++) {
    $sheet1->writeRow(['2020-1-' . $day, random_int(100, 999)]);
}

// make new sheet
$sheet2 = $excel->makeSheet('Feb');
$sheet2->setColFormat(1, '@date');
for ($day = 1; $day <= 29; $day++) {
    $sheet2->writeRow(['2020-2-' . $day, random_int(100, 999)]);
}

// make new sheet
$sheet3 = $excel->makeSheet('Mar');
$sheet3->setColFormat(1, '@date');
for ($day = 1; $day <= 31; $day++) {
    $sheet3->writeRow(['2020-3-' . $day, random_int(100, 999)]);
}

$sheet0 = $excel->getSheet('Total');

$sheet0->writeRow(['Jan 2020', '=SUM(Jan!B:B)']);
$sheet0->writeRow(['Feb 2020', '=SUM(Feb!B:B)']);
$sheet0->writeRow(['Mar 2020', '=SUM(Mar!B:B)']);

$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF