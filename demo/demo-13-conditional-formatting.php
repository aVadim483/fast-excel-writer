<?php

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use avadim\FastExcelWriter\Conditional\Conditional;
use \avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style;

$timer = microtime(true);

// Create Excel workbook
$excel = Excel::create();

$sheet = $excel->sheet();

$sheet->setColAutoWidth('a');

$sheet->nextRow();
$sheet->writeCell('cell: =5, >5');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::make('=', 5, [Style::FILL_COLOR => '#cfc']))
    ->applyConditionalFormatting(Conditional::make('>', 5, [Style::FILL_COLOR => '#fcc']))
;

$sheet->nextRow();
$sheet->writeCell('between [3, 6]');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::between([3, 6], [Style::FILL_COLOR => '#ccf']))
;

$sheet->nextRow();
$sheet->writeCell('not between [3, 6]');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::notBetween([3, 6], [Style::FILL_COLOR => '#ccf']))
;

$sheet->nextRow();
$sheet->writeCell('expression: MOD(RC,2)=0');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::expression('MOD(RC,2)=0', [Style::FILL_COLOR => '#ff9']))
;

$sheet->nextRow();
$sheet->writeCell('colorScale');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::colorScale('#f99', 'ff9', '9f9'))
;

$sheet->nextRow();
$sheet->writeCell('dataBar');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::dataBar('#99f'))
;

$sheet->nextRow();
$sheet->writeCell('aboveAverage');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::aboveAverage([Style::FILL_COLOR => '#9f9']))
;

$sheet->nextRow();
$sheet->writeCell('belowAverage');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::belowAverage([Style::FILL_COLOR => '#f99']))
;

$sheet->nextRow();
$sheet->writeCell('uniqueValues');
$sheet->writeCells(array_merge(range(0, 7), [7, 7]))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::uniqueValues([Style::FILL_COLOR => '#99f']))
;

$sheet->nextRow();
$sheet->writeCell('duplicateValues');
$sheet->writeCells(array_merge(range(0, 7), [7, 7]))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::duplicateValues([Style::FILL_COLOR => '#ff9']))
;

$sheet->nextRow();
$sheet->writeCell('top: 3');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::top(3, [Style::FILL_COLOR => '#f99']))
;

$sheet->nextRow();
$sheet->writeCell('topPercent: 10%');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::topPercent(10, [Style::FILL_COLOR => '#ff9']))
;

$sheet->nextRow();
$sheet->writeCell('lowPercent: 10%');
$sheet->writeCells(range(0, 9))
    ->applyOuterBorder(Style::BORDER_THIN)
    ->applyConditionalFormatting(Conditional::lowPercent(10, [Style::FILL_COLOR => '#ff9']))
;

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";