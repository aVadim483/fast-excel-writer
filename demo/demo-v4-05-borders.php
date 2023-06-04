<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;
use \avadim\FastExcelWriter\Style;

$timer = microtime(true);
$excel = Excel::create();
$sheet = $excel->getSheet();

$sheet->setColWidth([1, 2, 3, 4, 5, 6, 7], 4);
$sheet->setRowHeight([1, 2, 3, 4, 5], 20);

$area = $sheet->beginArea();

// simple border style
$style1 = [
    Style::BORDER => Style::BORDER_THIN
];

// border style with color
$style2 = [
    Style::BORDER => [
        Style::BORDER_ALL => [
            Style::BORDER_STYLE => Style::BORDER_THICK,
            Style::BORDER_COLOR => '#f00',
        ]
    ]
];

// extra border style
$style3 = [
    Style::BORDER => [
        Style::BORDER_TOP => Style::BORDER_NONE,
        Style::BORDER_LEFT => [
            Style::BORDER_STYLE => Style::BORDER_THICK,
            Style::BORDER_COLOR => '#f9009f',
        ],
        Style::BORDER_RIGHT => [
            Style::BORDER_STYLE => Style::BORDER_MEDIUM_DASH_DOT,
            Style::BORDER_COLOR => '#00f',
        ],
        Style::BORDER_BOTTOM => [
            Style::BORDER_STYLE => Style::BORDER_DOUBLE,
            Style::BORDER_COLOR => '#0ff',
        ],
    ]
];

$area->setStyle('B2', $style1);
$area->setStyle('D3', $style2);
$area->setStyle('F4', $style3);

$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF