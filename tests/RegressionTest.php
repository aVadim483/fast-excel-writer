<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Exceptions\Exception;
use avadim\FastExcelWriter\Writer\Writer;
use avadim\FastExcelReader\Excel as ExcelReader;
use PHPUnit\Framework\TestCase;

/**
 * Regression tests for the fixes of the 2026-07 audit
 */
final class RegressionTest extends TestCase
{
    protected array $savedFiles = [];

    protected function tearDown(): void
    {
        foreach ($this->savedFiles as $file) {
            if (file_exists($file)) {
                unlink($file);
            }
        }
        $this->savedFiles = [];
    }


    protected function saveCheckRead($excel, $testFileName): ExcelReader
    {
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }
        $this->savedFiles[] = $testFileName;
        $excel->save($testFileName);
        $this->assertTrue(file_exists($testFileName));
        $valid = ExcelReader::validate($testFileName, $errors);
        if ($errors) {
            $messages = [];
            foreach ($errors as $err) {
                $messages[] = $err->message . ' (' . $err->file . ')';
            }
            $this->fail('Invalid XLSX: ' . implode('; ', $messages));
        }
        $this->assertTrue($valid);

        return ExcelReader::open($testFileName);
    }


    protected function readXml(string $testFileName, string $entry): string
    {
        $zip = new ZipArchive();
        $zip->open($testFileName);
        $xml = $zip->getFromName($entry);
        $zip->close();

        return (string)$xml;
    }


    /**
     * The last valid row (1048576) and the last valid column (XFD) must be accepted,
     * anything beyond them must be rejected
     */
    public function testRowAndColumnLimits()
    {
        $testFileName = __DIR__ . '/regr_limits.xlsx';

        $excel = Excel::create(['Limits']);
        $sheet = $excel->sheet();

        $sheet->writeTo('XFD1', 'last col');
        $sheet->writeTo('A1048576', 'last row');

        $exception = null;
        try {
            $sheet->writeTo('A1048577', 'over the last row');
        }
        catch (Exception $e) {
            $exception = $e;
        }
        $this->assertNotNull($exception, 'Row number over ' . Excel::MAX_ROW . ' must be rejected');

        $reader = $this->saveCheckRead($excel, $testFileName);
        $cells = $reader->readCells();

        $this->assertEquals('last col', $cells['XFD1']);
        $this->assertEquals('last row', $cells['A1048576']);

        // the same check for the sequential writing API
        $excel = Excel::create(['Limits']);
        $sheet = $excel->sheet();
        $sheet->writeTo('XFC1', 'one before the last col');
        $sheet->writeCell('last col');

        $exception = null;
        try {
            $sheet->writeCell('over the last col');
        }
        catch (Exception $e) {
            $exception = $e;
        }
        $this->assertNotNull($exception, 'Column number over ' . Excel::MAX_COL . ' must be rejected');
    }


    /**
     * Leading and trailing spaces must survive the round trip via shared strings (xml:space="preserve")
     */
    public function testSpacesInSharedStrings()
    {
        $testFileName = __DIR__ . '/regr_spaces.xlsx';
        $values = ['  leading', 'trailing  ', '  both  ', ' ', "\ttab around\t"];

        foreach ([true, false] as $sharedString) {
            $excel = Excel::create(['Spaces'], ['shared_string' => $sharedString]);
            $sheet = $excel->sheet();
            $sheet->writeRow($values);

            $reader = $this->saveCheckRead($excel, $testFileName);
            $cells = $reader->readCells();

            $message = 'shared_string=' . var_export($sharedString, true);
            $this->assertEquals($values[0], $cells['A1'], $message);
            $this->assertEquals($values[1], $cells['B1'], $message);
            $this->assertEquals($values[2], $cells['C1'], $message);
            $this->assertEquals($values[3], $cells['D1'], $message);
            $this->assertEquals($values[4], $cells['E1'], $message);

            $xml = $this->readXml($testFileName, $sharedString ? 'xl/sharedStrings.xml' : 'xl/worksheets/sheet1.xml');
            $this->assertStringContainsString('xml:space="preserve"', $xml, $message);

            unlink($testFileName);
        }
    }


    /**
     * Formulas with escaped quotes ("") and with localized function names containing regex metacharacters
     */
    public function testFormulasWithQuotesAndLocalizedNames()
    {
        $testFileName = __DIR__ . '/regr_formulas.xlsx';

        $excel = Excel::create(['Formulas'], ['locale' => 'ru']);
        $sheet = $excel->sheet();
        $sheet->writeRow([1, 2]);
        // a string literal must not be touched by the formula converter, even when it contains
        // a localized function name or an escaped quote
        $sheet->writeTo('A2', '=IF(A1>0,"say ""СУММ"" here","no")');
        $sheet->writeTo('A3', '=СУММ(A1:B1)');
        $sheet->writeTo('A4', '=МНИМ.СУММ("1+2i","3+4i")');

        $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');

        // the literal keeps the localized name and both quotes
        $this->assertStringContainsString('<f>IF(A1&gt;0,&quot;say &quot;&quot;СУММ&quot;&quot; here&quot;,&quot;no&quot;)</f>', $xml);
        // function names outside of literals are converted
        $this->assertStringContainsString('<f>SUM(A1:B1)</f>', $xml);
        // a dot in a localized name is a regex metacharacter - it must be quoted, not matched as "any char"
        $this->assertStringContainsString('<f>IMSUM(&quot;1+2i&quot;,&quot;3+4i&quot;)</f>', $xml);
    }


    /**
     * An invalid UTF-8 byte must not drop the whole cell value
     */
    public function testInvalidUtf8()
    {
        $testFileName = __DIR__ . '/regr_utf8.xlsx';

        $excel = Excel::create(['Utf8']);
        $sheet = $excel->sheet();
        $sheet->writeRow(["begin\xB5end", "\xC3\x28 broken", 'valid текст']);

        $reader = $this->saveCheckRead($excel, $testFileName);
        $cells = $reader->readCells();

        $this->assertStringContainsString('begin', (string)$cells['A1']);
        $this->assertStringContainsString('end', (string)$cells['A1']);
        $this->assertStringContainsString('broken', (string)$cells['B1']);
        $this->assertEquals('valid текст', $cells['C1']);
    }


    /**
     * Numbers and dates must be written with a dot, whatever LC_NUMERIC is set (PHP 7.4 casts float by locale)
     */
    public function testCommaDecimalLocale()
    {
        $testFileName = __DIR__ . '/regr_locale.xlsx';
        $oldLocale = setlocale(LC_ALL, '0');
        $locale = setlocale(LC_ALL, 'de_DE.UTF-8', 'de_DE', 'de_DE@euro', 'German_Germany.1252', 'ru_RU.UTF-8', 'Russian_Russia.1251');

        try {
            if ($locale === false || strpos((string)sprintf('%.2f', 1.5), ',') === false) {
                $this->markTestSkipped('No comma decimal locale available on this system');
            }

            $excel = Excel::create(['Locale']);
            $sheet = $excel->sheet();
            $sheet->writeRow([3.14, -0.5, 1000000.125]);
            $sheet->writeTo('A2', '2026-08-16 12:30:45');
            $sheet->applyFormat('@DATETIME');

            $reader = $this->saveCheckRead($excel, $testFileName);
            $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');

            $this->assertStringNotContainsString(',', preg_replace('#<f>.*?</f>#s', '', $xml), 'No comma is allowed inside numeric values');
            $this->assertStringContainsString('<v>3.14</v>', $xml);
            $this->assertStringContainsString('<v>-0.5</v>', $xml);

            $cells = $reader->readCells();
            $this->assertEquals(3.14, $cells['A1']);
            $this->assertEquals(-0.5, $cells['B1']);
            $this->assertEquals(1000000.125, $cells['C1']);
        }
        finally {
            setlocale(LC_ALL, $oldLocale);
        }
    }


    /**
     * Two workbooks open at the same time must not share the state of the style manager
     */
    public function testTwoWorkbooksInOneProcess()
    {
        $testFileName1 = __DIR__ . '/regr_book1.xlsx';
        $testFileName2 = __DIR__ . '/regr_book2.xlsx';

        $excel1 = Excel::create(['Book1'], ['default_font' => ['font-name' => 'Arial', 'font-size' => 14]]);
        $excel2 = Excel::create(['Book2'], ['default_font' => ['font-name' => 'Courier New', 'font-size' => 9]]);

        $sheet1 = $excel1->sheet();
        $sheet2 = $excel2->sheet();

        // write into both workbooks alternately
        $sheet1->writeRow(['book one']);
        $sheet2->writeRow(['book two']);
        $sheet1->writeRow(['second row'], ['font' => ['style' => 'bold']]);
        $sheet2->writeRow(['second row'], ['font' => ['style' => 'bold']]);

        $reader1 = $this->saveCheckRead($excel1, $testFileName1);
        $reader2 = $this->saveCheckRead($excel2, $testFileName2);

        $this->assertEquals('book one', $reader1->readCells()['A1']);
        $this->assertEquals('book two', $reader2->readCells()['A1']);

        $styles1 = $this->readXml($testFileName1, 'xl/styles.xml');
        $styles2 = $this->readXml($testFileName2, 'xl/styles.xml');

        $this->assertStringContainsString('val="Arial"', $styles1);
        $this->assertStringNotContainsString('val="Courier New"', $styles1);

        $this->assertStringContainsString('val="Courier New"', $styles2);
        $this->assertStringNotContainsString('val="Arial"', $styles2);
    }


    /**
     * A sheet added after removeSheet() must not reuse the file name of the removed one
     */
    public function testRemoveSheetAndMakeSheet()
    {
        $testFileName = __DIR__ . '/regr_sheets.xlsx';

        $excel = Excel::create(['First', 'Second']);
        $excel->getSheet('First')->writeRow(['first']);
        $excel->getSheet('Second')->writeRow(['second']);

        $excel->removeSheet('Second');
        $third = $excel->makeSheet('Third');
        $third->writeRow(['third']);

        $reader = $this->saveCheckRead($excel, $testFileName);

        $this->assertEquals(['First', 'Third'], array_values($reader->getSheetNames()));
        $this->assertEquals('first', $reader->selectSheet('First')->readCells()['A1']);
        $this->assertEquals('third', $reader->selectSheet('Third')->readCells()['A1']);

        // each sheet must have its own xml file
        $workbookRels = $this->readXml($testFileName, 'xl/_rels/workbook.xml.rels');
        preg_match_all('#Target="(worksheets/sheet\d+\.xml)"#', $workbookRels, $m);
        $this->assertCount(2, $m[1]);
        $this->assertEquals($m[1], array_unique($m[1]), 'Sheet xml file names must be unique');
    }


    /**
     * Duplicate sheet names are not allowed (case-insensitively), the reserved name "History" either
     */
    public function testSheetNames()
    {
        $testFileName = __DIR__ . '/regr_names.xlsx';

        $this->assertEquals('History_', Writer::sanitizeSheetName('History'));
        $this->assertEquals('history_', Writer::sanitizeSheetName('history'));
        $this->assertEquals('History2', Writer::sanitizeSheetName('History2'));

        // different raw names sanitized to the same value must not produce two sheets with one name
        $excel = Excel::create(['History']);
        $excel->makeSheet('My:Sheet');
        $excel->makeSheet('My Sheet');
        $excel->makeSheet('My?Sheet');
        foreach ($excel->getSheets() as $sheet) {
            $sheet->writeRow(['x']);
        }

        $reader = $this->saveCheckRead($excel, $testFileName);
        $names = array_values($reader->getSheetNames());

        $this->assertCount(4, $names);
        $lowerNames = array_map('mb_strtolower', $names);
        $this->assertEquals($lowerNames, array_unique($lowerNames), 'Sheet names must be unique case-insensitively');
        $this->assertNotContains('history', $lowerNames, '"History" is reserved by Excel');
    }


    /**
     * Control characters must be encoded as _xHHHH_ (ST_Xstring) instead of being replaced with a space,
     * CR must be encoded too, a literal _xHHHH_ must be escaped as _x005F_xHHHH_
     */
    public function testControlCharactersAreEncoded()
    {
        $this->assertEquals('plain text', Writer::xmlEscapedString('plain text'));
        $this->assertEquals('line1_x000D_line2', Writer::xmlEscapedString("line1\rline2"));
        $this->assertEquals('bell_x0007_end', Writer::xmlEscapedString("bell\x07end"));
        $this->assertEquals('_x001F_', Writer::xmlEscapedString("\x1f"));
        // tab and new line are valid in XML and must be kept as is
        $this->assertEquals("keep\ttab\nand nl", Writer::xmlEscapedString("keep\ttab\nand nl"));
        // a literal "_xHHHH_" must not be read back as a control character
        $this->assertEquals('_x005F_x000D_ literal', Writer::xmlEscapedString('_x000D_ literal'));
        $this->assertEquals('a &amp; &lt;b&gt;', Writer::xmlEscapedString('a & <b>'));
        // DEL is valid in XML, but it would be replaced with a space, so it is encoded as well
        $this->assertEquals('del_x007F_end', Writer::xmlEscapedString("del\x7fend"));

        $testFileName = __DIR__ . '/regr_control_chars.xlsx';
        $values = ["line1\rline2", "bell\x07end", '_x000D_ literal', "keep\ttab\nnl", "del\x7fend"];

        foreach ([true, false] as $sharedString) {
            $excel = Excel::create(['Chars'], ['shared_string' => $sharedString]);
            $sheet = $excel->sheet();
            $sheet->writeRow($values);
            $sheet->writeTo('A2', ['=A1&"x"', "result\rvalue"]);

            $reader = $this->saveCheckRead($excel, $testFileName);
            $cells = $reader->readCells();
            $message = 'shared_string=' . var_export($sharedString, true);

            // the round trip is complete: the reader decodes _xHHHH_ back to the character (as Excel does)
            $this->assertEquals("line1\rline2", $cells['A1'], $message);
            $this->assertEquals("bell\x07end", $cells['B1'], $message);
            $this->assertEquals("keep\ttab\nnl", $cells['D1'], $message);
            $this->assertEquals("del\x7fend", $cells['E1'], $message);
            // a literal _xHHHH_ is not decoded twice, it survives as the text it was
            $this->assertEquals('_x000D_ literal', $cells['C1'], $message);

            $xml = $this->readXml($testFileName, $sharedString ? 'xl/sharedStrings.xml' : 'xl/worksheets/sheet1.xml');
            $this->assertStringContainsString('line1_x000D_line2', $xml, $message);
            $this->assertStringContainsString('_x005F_x000D_ literal', $xml, $message);

            // a pre-calculated formula result is escaped the same way
            $sheetXml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');
            $this->assertStringContainsString('<v>result_x000D_value</v>', $sheetXml, $message);

            unlink($testFileName);
        }
    }
}
