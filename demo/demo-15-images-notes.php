<?php

require_once __DIR__ . '/../vendor/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\ImageStyle;

$imageFile = __DIR__ . '/logo/excel-logo.png';

$timer = microtime(true);

// Create Excel workbook
$excel = Excel::create(['Images & Notes']);
$sheet = $excel->sheet();

$sheet->setColWidths([24, 24, 24]);
$sheet->setRowHeight(2, 60);

$sheet->writeRow(['Image as is', 'Style as array', 'Style as ImageStyle'], ['font-style' => 'bold', 'text-align' => 'center']);

// Insert an image to the cell from a local file (URL or base64 string are also supported)
$sheet->addImage('A2', $imageFile);

// Set image size and offset via an array of options
$sheet->addImage('B2', $imageFile, ['height' => 40, 'x' => 5, 'y' => 5]);

// The same result with the ImageStyle class and its fluent interface
$imageStyle = (new ImageStyle())
    ->height(40)
    ->offset(5, 5)
    ->hyperlink('https://github.com/aVadim483/fast-excel-writer');
$sheet->addImage('C2', $imageFile, $imageStyle);

// You can add notes to any cells...
$sheet->writeTo('A4', 'Cell with a note')
    ->addNote('This is a note for cell A4');

// ...change their size and color...
$sheet->writeTo('B4', 'Note with custom style')
    ->addNote("Line 1\nLine 2", ['width' => 200, 'height' => 100, 'fill_color' => '#ffcccc']);

// ...and make them always visible
$sheet->addNote('C4', 'This note is always visible', ['show' => true]);

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF
