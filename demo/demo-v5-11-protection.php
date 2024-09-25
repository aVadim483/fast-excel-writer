<?php

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style;

$timer = microtime(true);

// Create Excel workbook
$excel = Excel::create(['Without Psw', 'With Psw']);

$sheet = $excel->sheet();
$sheet->setColWidths(['B' => 36, 'C' => 12]);

$sheet->writeTo('B3:C3', 'Calculate Mortgage Payments')
    ->applyBgColor('#cfd')
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyFont('Arial', 14, 'bold')
    ->applyTextAlign('center', 'center')
;

$sheet->writeTo('B4:C4', 'Enter info in yellow cells')
    ->applyTextAlign('center', 'center');

$sheet->setTopLeftCell('B8');
$style = [
    'bg_color' => '#fffccc',
    'border' => 'thin',
];

$sheet->writeCell('Principal');
$sheet->writeCell(100000, ['format' => '@money'])
    ->applyStyle($style)
    ->applyUnlock()
    ->applyNamedRange('Principal');
$sheet->nextRow();

$sheet->writeCell('Annual interest rate');
$sheet->writeCell(0.05, ['format' => '@percent'])
    ->applyStyle($style)
    ->applyUnlock()
    ->applyNamedRange('Rate');
$sheet->nextRow();

$sheet->writeCell('Number of monthly payments');
$sheet->writeCell(360)
    ->applyStyle($style)
    ->applyUnlock()
    ->applyNamedRange('Months');
$sheet->skipRow(2);

$sheet->writeCell('Monthly mortgage payment')->applyFontStyleBold();
$sheet->writeCell('=PMT((Rate/12),Months,-Principal,0)', ['format' => '@money'])
    ->applyFontStyleBold()
    ->applyHide();

$sheet->protect();

$sheet = $excel->sheet(2);
$area = $sheet->beginArea();
$area->writeTo('C2:F2', 'You can write to yellow area only')->applyTextAlign('center', 'center');

$area = $sheet->makeArea('C4:F10')->applyOuterBorder('thick')->applyBgColor('#ccc');
$area = $sheet->makeArea('d5:e9')->applyOuterBorder('thin')->applyBgColor('#fffccc')->applyUnlock();
$sheet->protect('qwerty');

$excel->protect();

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF