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
        Style::FONT_STYLE => [Style::FONT_STYLE_BOLD, Style::FONT_STYLE_ITALIC],
    ],
    Style::TEXT_ALIGN => Style::TEXT_ALIGN_CENTER,
    Style::BORDER => Style::BORDER_THICK,
];

$sheet->writeRow($row, $rowOptions);
$cnt1 = $sheet->rowCountWritten + 1;

$max = 20;
for ($i = 0; $i < $max; $i++) {
    $row = [$i, random_int(100, 999), '=RC[-1]*0.1'];
    $sheet->writeRow($row);
}

$cnt2 = $sheet->rowCountWritten;
$totals = [
    ['Total', "=SUM(B$cnt1:B$cnt2)", ''],
    ['Average', '', "=ROUND(AVERAGE(C$cnt1:C$cnt2),1)"]
];
$rowOptions = [
    Style::FONT => [
        Style::FONT_STYLE => Style::FONT_STYLE_BOLD,
    ],
    Style::BORDER => [Style::BORDER_TOP => Style::BORDER_DOUBLE],
];

$sheet->writeRow($totals[0], $rowOptions);
$sheet->writeRow($totals[1])
    ->applyFontStyleBold()
    ->applyBorderTop(Style::BORDER_DOUBLE);

$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF