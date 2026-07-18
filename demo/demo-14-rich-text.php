<?php

require_once __DIR__ . '/../vendor/autoload.php';

$outFileName = __DIR__ . '/output/' . basename(__FILE__, '.php') . '.xlsx';

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\RichText\RichText;

$timer = microtime(true);

// Create Excel workbook
$excel = Excel::create(['RichText Demo']);
$sheet = $excel->sheet();

$sheet->setColWidths([44, 60]);

// The first way - add fragments and set their styles via fluent interface
$richText = new RichText();
$richText->addText('ATTENTION!')->setBold();
$richText->addText(' The product is reserved for ');
$richText->addText('5 days')->setUnderline()->setColor('red');

$sheet->writeRow(['Fragments via fluent interface', $richText]);

// The second way - pass fragments to the constructor and style them by index
$richText = new RichText('ATTENTION! ', 'The product is reserved for ', '5 days');
$richText->fragment(0)->setBold();
$richText->fragment(2)->setUnderline()->setColor('f00');

$sheet->writeRow(['Fragments by index', $richText]);

// The third way - use simple tags
$richText = new RichText('<b>ATTENTION!</b> The product is reserved for <u><c=red>5 days</c></u>');

$sheet->writeRow(['Simple tags', $richText]);

// Different fonts, sizes and colors in one cell
$richText = new RichText();
$richText->addText('Arial')->setFont('Arial');
$richText->addText(' Courier New')->setFont('Courier New')->setColor('#00a000');
$richText->addText(' big')->setSize(15);
$richText->addText(' bigger')->setSize(20)->setItalic();

$sheet->writeRow(['Fonts, sizes and colors', $richText]);

// Also, you can use rich text in notes
$sheet->writeRow(['Rich text in the note (hover the cell)', 'Cell with a note'])
    ->addNote('B5', new RichText('here is <c=f00>red</c> and <c=00f>blue</c> text'));

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF
