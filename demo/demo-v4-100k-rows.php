<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;

$chars = 'abcdefghijklmnopqrstuvwxyz0123456789 ';
$s = '';
for ($j = 0; $j < 16192; $j++) {
    $s .= $chars[mt_rand() % 36];
}

$countK = $_GET['count'] ? (int)$_GET['count'] : 100;
$rowCount = $countK * 1024;

$timer = microtime(true);
$excel = Excel::create([$countK . 'K']);
$sheet = $excel->getSheet();

$sheet->setColFormats(['@integer', '@string', '0.00', '@string', '@string']);

$n = 0;
for($i = 0; $i < $rowCount; $i++) {
    $s1 = substr($s, mt_rand() % 400, mt_rand() % 5 + 5);
    $s2 = substr($s, mt_rand() % 800, mt_rand() % 5 + 5);
    $sheet->writeRow([$i, $s1, ($i % 100) / 100, $chars[$n], $s2]);
    if (!$chars[++$n]) {
        $n = 0;
    }
}

$excel->save($outFileName);

$time = round(microtime(true) - $timer, 3);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', $time, ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";
echo 'total rows: ', $rowCount, "<br>\n";
echo 'speed: ', round($rowCount / $time, 3), " rows/sec<br>\n";

// EOF