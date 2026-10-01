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

// Scientific notation with subscript and superscript
$sheet->writeRow(['Subscript and superscript tags', new RichText('H<sub>2</sub>O, m<sup>2</sup>')]);
$scientific = new RichText();
$scientific->addText('CO')->addText('2')->setSubscript();
$scientific->addText(' and E = mc')->addText('2')->setSuperscript()->setBold();
$sheet->writeRow(['Subscript and superscript methods', $scientific]);

// Strikethrough and underline
$decorations = new RichText();
$decorations->addText('Old price')->setStrike();
$decorations->addText(' New price')->setUnderline(true);
$sheet->writeRow(['Strikethrough and double underline', $decorations]);
$sheet->writeRow(['Strikethrough tags', new RichText('<strike>old</strike> <del>removed</del> new')]);

// Literal text is escaped automatically; fractional sizes are preserved.
$literal = new RichText();
$literal->addText('A & B, x < y')->setSize(10.5)->setBold();
$literal->addText(' regular')->setBold(false)->setItalic(false)->setStrike(false)->removeUnderline();
$sheet->writeRow(['Escaping, fractional size and explicit formatting', $literal]);

// Save to XLSX-file
$excel->save($outFileName);

echo '<b>', basename(__FILE__, '.php'), "</b><br>\n<br>\n";
echo 'out filename: ', $outFileName, "<br>\n";
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec', "<br>\n";
echo 'memory peak usage: ', memory_get_peak_usage(true), "<br>\n";

// EOF
