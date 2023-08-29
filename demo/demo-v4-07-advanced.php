<?php
include_once __DIR__ . '/../src/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use \avadim\FastExcelWriter\Excel;
use \avadim\FastExcelWriter\Style;

// *****************
// PREPARE DEMO DATA
$demoNames = [
    ['John', 'Evan', 'Giovanni', 'Janusz', 'Hans', 'Johann', 'Jean'. 'Peter', 'Pedro', 'Pierre', 'Pietro', 'Francesco', 'James', 'Mateusz', 'Miguel'],
    ['Smith', 'Johnson', 'Smirnov', 'Lee', 'Wong', 'Muller', 'Schmidt', 'Silva', 'Santos', 'Andersson', 'Johansson', 'Russo', 'Kowalski', 'Novak'],
];
$lorem = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua';
$loremLen = strlen($lorem);
$time1 = strtotime('1940-1-1');
$time2 = strtotime('1999-12-31');
$data = [];
$rowCount = 100;
for ($i = 1; $i <= $rowCount; $i++) {
    $row = [
        $i, // #
        $demoNames[0][array_rand($demoNames[0])], // f.name
        $demoNames[1][array_rand($demoNames[1])], // l.name
        date('Y-m-d', random_int($time1, $time2)), // b.day
        '=ROUNDDOWN((TODAY()-RC[-1])/365,0)', // age
        random_int(100, 200) / 10, // quant.
        random_int(1000, 9000) / 10, // price
        '=RC[-1]*RC[-2]',
        random_int(1, 3) * 5 / 100,
        '=RC[-1]*RC[-2]',
        substr($lorem, 0, random_int(11, $loremLen)),
    ];
    $data[] = $row;
}
$colors = ['cc9', 'c9c', 'cc9', 'c9f', '9cf', '9fc', '36f', '3f6', '63f', '36c', '3c6', '63c'];
$images = glob(__DIR__ . '/logo/excel-logo.*');
$noteColors = ['', '#f8d7e9', '#bad6f1', '#f9efe7', '#d7eff7'];
// *****************


$timer = microtime(true);

// Create new Excel book
$excel = Excel::create(['Demo']);

// Set locale - In most cases, the locale is automatically set correctly,
// but sometimes you need to do it manually
$excel->setLocale('ru');

// Get the first sheet
$sheet = $excel->getSheet();
$sheet
    ->pageOrientationLandscape()  // set page orientation
    ->pageFitToWidth(1)  // fit width to 1 page
    ->pageFitToHeight(1);// fit height to 1 page

$headerStyle = [
    'font' => [
        'size' => 24,
        'style' => 'bold'
    ],
    'text-align' => 'center',
    'vertical-align' => 'center',
];

/*  DOCUMENT HEADER */

$cells = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1'];
foreach($cells as $cell) {
    $color = '#' . $colors[array_rand($colors)] . $colors[array_rand($colors)];
    // set background colors for specified cells
    $sheet->setBgColor($cell, $color);
}
$sheet->writeTo('a2', '');
foreach ($images as $n => $image) {
    $cell = Excel::cellAddress(3, $n + 2);
    $sheet->addImage($cell, $image, ['height' => 40]);
    $bgColor = $noteColors[$n];
    $sheet->addNote($cell, basename($image), ['bg_color' => $bgColor]);
}
$sheet->setRowHeight(3, 50);

// Begin an area for direct write
$area = $sheet->beginArea();

// Write value to automerged cells
$area->setValue('A4:K4', 'This is demo XLSX-sheet')
    ->applyFontStyleBold()
    ->applyFontSize(24)
    ->applyTextCenter();

$area->setValue('E5:I5', 'avadim/fast-excel-writer', ['hyperlink' => 'https://github.com/aVadim483/fast-excel-writer', 'align'=>'center']);

$area
    ->setValue('J6', 'Date:', ['text-align' => 'right'])
    ->setValue('K6', date('Y-m-d H:i:s'), ['font-style' => 'bold', 'format' => '@datetime', 'text-align' => 'left'])
;

/* TABLE HEADER */

// Begin new area (specify left top cell)
$area = $sheet->beginArea('A8');
//var_dump($area->getBeginAddress()); exit;
// You can use R1C1-notation, start position in A6
$area
    ->setValue('RC:R[1]C', '#') // Merge vertical cells
    ->setValue('RC1:RC2', 'People') // Merge horizontal cells
    ->setValue('R1C1', 'First Name') // Single cell
    ->setValue('R1C2', 'Last Name')
    ->setValue('RC3:R1C3', 'Birthday')
    ->setValue('RC4:R1C4', 'Age')
    ->setValue('RC5:R1C5', 'Quantity')
    ->setValue('RC6:R1C6', 'Price')
    ->setValue('RC7:R1C7', 'Cost')
    ->setValue('RC8:R1C8', 'Tax Rate')
    ->setValue('RC9:R1C9', 'Tax Value')
    ->setValue('RC10:R1C10', 'Description')
;

$tableHeaderStyle = [
    'font' => ['style' => 'bold'],
    'fill' => '#eee',
    'text-align' => 'center',
    'vertical-align' => 'center',
    'border' => 'thin',
];

$area->setStyle('RC:R1C10', $tableHeaderStyle);
$area->setOuterBorder('R0C0:R1C10', Style::BORDER_THICK);

$sheet->writeAreas();

/*
 * You can set three levels for cell style^
 * 1. Default style for sheet via setDefaultStyle()
 * 2. Style for column via setColStyle()
 * 3. Style for cells in row via writeRow($row, $rowOptions)
 */

// Default style options for all next cells
$sheet->setDefaultStyle(['vertical-align' => 'top']);

// Set widths of columns from the first (A)
$sheet->setColWidths([5, 16, 16, 'auto']);

// Set width of the column
$sheet->setColWidth(['G', 'H', 'J'], 14);

// Set formats of columns from the first (A); null - default format
$sheet->setColFormats([null, '@', '@', '@date', '0', '0.00', '@money', '@money']);

// Set style and width for specified column
$sheet->setColOptions('K', ['text-wrap' => true, 'width' => 32]);

// Set options for specified columns in the row
$cellStyles = ['I' => ['format' => '@percent'], 'j' => ['format' => '@money']];
foreach($data as $n => $row) {
    if ($n % 2) {
        $rowOptions = ['fill' => '#eee'];
    }
    else {
        $rowOptions = null;
    }
    $sheet->writeRow($row, $rowOptions, $cellStyles);
}

$totalRow = [];
$sheet->writeRow($totalRow, ['font' => 'bold', 'border-top' => 'double']);

$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF