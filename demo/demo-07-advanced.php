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
        $i,
        $demoNames[0][array_rand($demoNames[0])],
        $demoNames[1][array_rand($demoNames[1])],
        date('Y-m-d', random_int($time1, $time2)),
        '=ROUNDDOWN((TODAY()-RC[-1])/365,0)',
        //'=RC[-1]',
        random_int(100, 200) / 10,
        random_int(1000, 9000) / 10,
        '=RC[-1]*RC[-2]',
        random_int(1, 3) * 5 / 100,
        substr($lorem, 0, random_int(11, $loremLen)),
    ];
    $data[] = $row;
}
$colors = ['cc9', 'c9c', 'cc9', 'c9f', '9cf', '9fc', '36f', '3f6', '63f', '36c', '3c6', '63c'];
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
// Begin an area for direct write
$area = $sheet->beginArea();

$cells = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1'];
foreach($cells as $cell) {
    $color = '#' . $colors[array_rand($colors)] . $colors[array_rand($colors)];
    // set background colors for specified cells
    $area->setBgColor($cell, $color);
}

// Write value to automerged cells
$area->setValue('A2:J2', 'This is demo XLSX-sheet', $headerStyle);

$area
    ->setValue('H4', 'Date', ['text-align' => 'right'])
    ->setValue('I4:J4', date('Y-m-d H:i:s'), ['font' => 'bold', 'format' => 'datetime', 'text-align' => 'left'])
;

/* TABLE HEADER */
// Begin new area (specify left top cell)
$area = $sheet->beginArea('A6');

// You can use R1C1-notation,

$area
    ->setValue('RC:R[1]C', '#')
    ->setValue('RC1:RC2', 'People')
    ->setValue('R1C1', 'First Name')
    ->setValue('R1C2', 'Last Name')
    ->setValue('RC3:R1C3', 'Birthday')
    ->setValue('RC4:R1C4', 'Age')
    ->setValue('RC5:R1C5', 'Quantity')
    ->setValue('RC6:R1C6', 'Price')
    ->setValue('RC7:R1C7', 'Cost')
    ->setValue('RC8:R1C8', 'Tax')
    ->setValue('RC9:R1C9', 'Description')
;

$tableHeaderStyle = [
    'font' => ['style' => 'bold'],
    'fill' => '#eee',
    'text-align' => 'center',
    'vertical-align' => 'center',
    'border' => 'thin',
];

$area->setStyle('RC:R1C9', $tableHeaderStyle);
$area->setOuterBorder('R0C0:R1C9', Style::BORDER_THICK);

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
$sheet->setColWidths([5, 16, 16, 13]);

// Set width of the column
$sheet->setColWidth(['G', 'H'], 14);

// Set formats of columns from the first (A); null - default format
$sheet->setColFormats([null, '@', '@', 'date', '0', '0.00', 'money', 'money']);

// Set width for specified column
$sheet->setColWidth('J', 20);

// Set style for specified column
$sheet->setColStyle('J', ['text-wrap' => true]);

// Set options for specified columns in the row
$rowOptions = ['I' => ['format' => 'percent']];
foreach($data as $row) {
    $sheet->writeRow($row, $rowOptions);
}

$excel->save($outFileName);

echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF