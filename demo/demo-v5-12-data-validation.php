<?php

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use avadim\FastExcelWriter\DataValidation\DataValidation;
use \avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style;

$timer = microtime(true);

// Create Excel workbook
$excel = Excel::create(['Demo', 'Data']);

$sheet = $excel->sheet('Data');
// Write data for dropdown
$sheet->writeRow(['ddd', 'eee', 'fff']);
// Write named range for dropdown
$sheet->writeRow(['item3-1', 'item3-2', 'item3-3'])->applyNamedRange('list');

$sheet = $excel->sheet('Demo');
$sheet->setColWidthAuto('A');

$style = [
    Style::BORDER => Style::BORDER_THICK,
    Style::FILL_COLOR => '#fffccc',
];

$validation2 = DataValidation::dropDown('=data!A1:C1');
$validation3 = DataValidation::dropDown('=data!list');



$validation = DataValidation::integer('>', 10)
    ->setPrompt('Number greater than 10')
    ->setError('Please enter correct number greater than 10');
;
$sheet->writeCell('Decimal 1');
$sheet->nextCell()
    ->applyStyle($style)
    ->applyDataValidation($validation)
;

$sheet->skipRow();
$validation = DataValidation::dropDown(['aaa', 'bbb', 'ccc']);
$sheet->writeCell('DropDown List 1');
$sheet->writeCell('')
    ->applyStyle($style)
    ->applyDataValidation($validation)
;

$sheet->nextRow();
$sheet->writeCell('DropDown List 2');
$sheet->writeCell('')
    ->applyBorder(Style::BORDER_THICK)
    ->applyBgColor('#fffccc')
    ->applyDataValidation($validation2)
;

$sheet->nextRow();
$sheet->writeCell('DropDown List 3');
$sheet->writeCell('')
    ->applyBorder(Style::BORDER_THICK)
    ->applyBgColor('#fffccc')
    ->applyDataValidation($validation3)
;

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF