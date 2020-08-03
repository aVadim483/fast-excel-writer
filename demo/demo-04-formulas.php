<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;
use \avadim\FastExcelWriter\Style;

$timer = microtime(true);
$excel = Excel::create(['Sheet1']);
$sheet = $excel->getSheet();

$row = ['#', 'Number', '\=RC[-1]*0.1'];
$rowOptions = [
    Style::FONT => [
        //Style::FONT_SIZE => 14,
        Style::FONT_STYLE => Style::FONT_STYLE_BOLD,
    ],
    Style::TEXT_ALIGN => Style::TEXT_ALIGN_CENTER,
    Style::BORDER => Style::BORDER_THICK,
];

$sheet->writeRow($row, $rowOptions);
$cnt1 = $sheet->rowCount + 1;

for ($i = 0; $i < 20; $i++) {
    $row = [$i, random_int(100, 999), '=RC[-1]*0.1'];
    $sheet->writeRow($row);
}

$cnt2 = $sheet->rowCount;
$row = ['Total', "=SUM(B$cnt1:B$cnt2)", "=AVERAGE(C$cnt1:C$cnt2)"];
$rowOptions = [
    Style::FONT => [
        Style::FONT_STYLE => Style::FONT_STYLE_BOLD,
    ],
    Style::BORDER => [Style::BORDER_TOP => Style::BORDER_DOUBLE],
];

$sheet->writeRow($row, $rowOptions);

$excel->save($outFileName);

echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF