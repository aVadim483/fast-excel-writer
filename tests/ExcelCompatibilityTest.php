<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Options;
use avadim\FastExcelWriter\Sheet;
use avadim\FastExcelWriter\Exceptions\Exception;
use avadim\FastExcelWriter\Conditional\Conditional;
use avadim\FastExcelWriter\DataValidation\DataValidation;
use avadim\FastExcelWriter\Writer\Writer;
use avadim\FastExcelReader\Excel as ExcelReader;
use PHPUnit\Framework\TestCase;

/**
 * Tests of the compatibility with Excel: things Excel writes in its own way
 */
final class ExcelCompatibilityTest extends TestCase
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
     * Data validation referring to another sheet must go to the x14 extension list,
     * exactly as Excel itself writes it
     */
    public function testExternalDataValidation()
    {
        $testFileName = __DIR__ . '/compat_validation.xlsx';

        $excel = Excel::create(['Main', 'Lists']);
        $lists = $excel->getSheet('Lists');
        $lists->writeRow(['red']);
        $lists->writeRow(['green']);

        $sheet = $excel->getSheet('Main');
        $sheet->writeRow(['pick one']);
        $sheet->addDataValidation('A2:A10', DataValidation::list('=Lists!$A$1:$A$2'));
        $sheet->addDataValidation('B2:B10', DataValidation::list(['one', 'two']));

        $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');

        // the local rule stays in the plain element
        $this->assertStringContainsString('<dataValidations count="1">', $xml);
        $this->assertStringContainsString('<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B2:B10">', $xml);
        // the rule referring to another sheet goes to the extension list and is not duplicated
        $this->assertStringNotContainsString('sqref="A2:A10"', $xml);
        $this->assertStringContainsString('<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"', $xml);
        $this->assertStringContainsString('<x14:formula1><xm:f>Lists!$A$1:$A$2</xm:f></x14:formula1>', $xml);
        $this->assertStringContainsString('<xm:sqref>A2:A10</xm:sqref>', $xml);
        // <extLst> must be the last element of the worksheet
        $this->assertStringEndsWith('</extLst></worksheet>', trim($xml));
    }


    /**
     * Titles and messages of a data validation are free text - without escaping the file is broken
     */
    public function testDataValidationEscaping()
    {
        $testFileName = __DIR__ . '/compat_validation_esc.xlsx';

        $excel = Excel::create(['Main']);
        $sheet = $excel->sheet();
        $sheet->writeRow(['value']);
        $validation = DataValidation::integer('between', [1, 10])
            ->setError('Value must be "1" & 10 <inclusive>', 'Bad & "value"')
            ->setPrompt('Enter 1 & 10', 'Hint <here>');
        $sheet->addDataValidation('A2:A10', $validation);

        // the file must be readable, that is the whole point
        $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');

        $this->assertStringContainsString('errorTitle="Bad &amp; &quot;value&quot;"', $xml);
        $this->assertStringContainsString('error="Value must be &quot;1&quot; &amp; 10 &lt;inclusive&gt;"', $xml);
        $this->assertStringContainsString('promptTitle="Hint &lt;here&gt;"', $xml);
        $this->assertStringContainsString('prompt="Enter 1 &amp; 10"', $xml);
    }


    /**
     * Conditional formatting referring to another sheet must go to the x14 extension list
     * with the style inlined into the rule
     */
    public function testExternalConditionalFormatting()
    {
        $testFileName = __DIR__ . '/compat_conditional.xlsx';

        $excel = Excel::create(['Main', 'Limits']);
        $excel->getSheet('Limits')->writeRow([100]);

        $sheet = $excel->getSheet('Main');
        $sheet->writeRow([1, 2]);
        $sheet->addConditionalFormatting('A2:A10', Conditional::greaterThan('=Limits!$A$1')->setFillColor('#FF0000'));
        $sheet->addConditionalFormatting('B2:B10', Conditional::greaterThan(5)->setFillColor('#00FF00'));

        $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');

        // the local rule stays in the plain element
        $this->assertStringContainsString('<conditionalFormatting sqref="B2:B10">', $xml);
        $this->assertStringNotContainsString('<conditionalFormatting sqref="A2:A10">', $xml);
        // the external one is in the extension list with an inline <x14:dxf>
        $this->assertStringContainsString('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"', $xml);
        $this->assertStringContainsString('<xm:f>Limits!$A$1</xm:f>', $xml);
        $this->assertStringContainsString('<xm:sqref>A2:A10</xm:sqref>', $xml);
        $this->assertMatchesRegularExpression('#<x14:dxf>.*<fill>.*FFFF0000.*</fill></x14:dxf>#', $xml);
        $this->assertStringNotContainsString('<x14:dxf><dxf>', $xml);
        // priorities of both rules are still unique
        preg_match_all('#priority="(\d+)"#', $xml, $m);
        $this->assertEquals($m[1], array_unique($m[1]));
    }


    /**
     * Default formats of the workbook override formats of the locale
     */
    public function testDefaultFormats()
    {
        $testFileName = __DIR__ . '/compat_formats.xlsx';

        // via options of the workbook
        $excel = Excel::create(['Dates'], [
            'default_date_format' => 'DD.MM.YYYY',
            'default_datetime_format' => 'DD.MM.YYYY HH:MM',
        ]);
        $sheet = $excel->sheet();
        $sheet->writeRow(['2026-08-16'], ['format' => '@date']);
        $sheet->writeRow([new \DateTime('2026-08-16 10:20:30')]);

        $this->saveCheckRead($excel, $testFileName);
        $styles = $this->readXml($testFileName, 'xl/styles.xml');
        $this->assertStringContainsString('formatCode="DD.MM.YYYY"', $styles);
        $this->assertStringContainsString('formatCode="DD.MM.YYYY HH:MM"', $styles);
        unlink($testFileName);

        // via the Options object
        $options = Options::create()->defaultDateFormat('YYYY/MM/DD');
        $excel = Excel::create(['Dates'], $options);
        $excel->sheet()->writeRow(['2026-08-16'], ['format' => '@date']);
        $this->saveCheckRead($excel, $testFileName);
        $this->assertStringContainsString('formatCode="YYYY/MM/DD"', $this->readXml($testFileName, 'xl/styles.xml'));
        unlink($testFileName);

        // a format of the workbook is stronger than a format of the locale, even if the locale is set later
        $excel = Excel::create(['Dates']);
        $excel->setDefaultDateFormat('YYYY/MM/DD');
        $excel->setDefaultTimeFormat('HH:MM');
        $excel->setLocale('de');
        $sheet = $excel->sheet();
        $sheet->writeRow(['2026-08-16'], ['format' => '@date']);
        $sheet->writeRow(['10:20:30'], ['format' => '@time']);

        $this->saveCheckRead($excel, $testFileName);
        $styles = $this->readXml($testFileName, 'xl/styles.xml');
        $this->assertStringContainsString('formatCode="YYYY/MM/DD"', $styles);
        $this->assertStringContainsString('formatCode="HH:MM"', $styles);
        // the German date format of the locale is overridden
        $this->assertStringNotContainsString('formatCode="DD.MM.YYYY"', $styles);
        unlink($testFileName);

        // the locale alone still works
        $excel = Excel::create(['Dates'], ['locale' => 'de']);
        $excel->sheet()->writeRow(['2026-08-16'], ['format' => '@date']);
        $this->saveCheckRead($excel, $testFileName);
        $this->assertStringContainsString('formatCode="DD.MM.YYYY"', $this->readXml($testFileName, 'xl/styles.xml'));
    }


    /**
     * Serial numbers of the 1900 date system, including the day Excel invented (29.02.1900)
     */
    public function testDateSerialNumbers()
    {
        $expected = [
            '1899-12-31' => 0,
            '1900-01-01' => 1,
            '1900-02-28' => 59,
            '1900-02-29' => 60, // does not exist in the real calendar, but exists in Excel
            '1900-03-01' => 61,
            '1900-12-31' => 366,
            '1901-01-01' => 367,
            '1904-01-01' => 1462,
            '2000-02-29' => 36585,
            '2026-08-16' => 46250,
            '9999-12-31' => 2958465,
        ];
        foreach ($expected as $date => $serial) {
            $this->assertEqualsWithDelta($serial, Writer::convertDateTime($date), 1e-9, 'Wrong serial number of ' . $date);
        }

        // the fake leap day keeps the time part too
        $this->assertEqualsWithDelta(60.5, Writer::convertDateTime('1900-02-29 12:00:00'), 1e-9);
        $this->assertEqualsWithDelta(1.5, Writer::convertDateTime('1900-01-01 12:00:00'), 1e-9);

        // dates before the epoch of Excel are rejected
        $this->assertFalse(Writer::convertDateTime('1899-12-30'));
        $this->assertFalse(Writer::convertDateTime('1898-05-05'));

        // and the same in a real file
        $testFileName = __DIR__ . '/compat_dates.xlsx';
        $excel = Excel::create(['Dates']);
        $sheet = $excel->sheet();
        $sheet->writeRow(['1900-02-28', '1900-02-29', '1900-03-01'], ['format' => '@date']);

        $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<v>59</v>', $xml);
        $this->assertStringContainsString('<v>60</v>', $xml);
        $this->assertStringContainsString('<v>61</v>', $xml);
    }


    /**
     * A numeric string longer than 15 significant digits (an ID, a barcode) must stay a text,
     * otherwise Excel silently rounds it
     */
    public function testLongNumbersAreKeptAsText()
    {
        $testFileName = __DIR__ . '/compat_long_numbers.xlsx';

        $excel = Excel::create(['Numbers'], ['auto_convert_number' => true]);
        $sheet = $excel->sheet();
        $sheet->writeRow(['12345678901234567890', '4006381333931', '123', '1.5', '007', '000123456789012345678']);

        $reader = $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');
        $cells = $reader->readCells();

        // 20 digits - text
        $this->assertStringContainsString('<t xml:space="preserve">12345678901234567890</t>', $xml);
        $this->assertSame('12345678901234567890', $cells['A1']);
        // leading zeros do not count as significant digits, but the value is a text anyway
        $this->assertSame('000123456789012345678', $cells['F1']);
        // 13 digits - a normal number
        $this->assertStringContainsString('<v>4006381333931</v>', $xml);
        $this->assertEquals(4006381333931, $cells['B1']);
        $this->assertEquals(123, $cells['C1']);
        $this->assertEquals(1.5, $cells['D1']);
        // a number with a leading zero is a text as before
        $this->assertSame('007', $cells['E1']);
    }


    /**
     * A value longer than 32767 characters makes the whole file unreadable for Excel, so it is truncated
     */
    public function testCellLengthLimit()
    {
        $testFileName = __DIR__ . '/compat_long_text.xlsx';
        $limit = Writer::MAX_CELL_LENGTH;

        $excel = Excel::create(['Long'], ['shared_string' => false]);
        $sheet = $excel->sheet();
        $sheet->writeRow([str_repeat('a', $limit + 5000), str_repeat('я', $limit + 5000), str_repeat('b', $limit), 'short']);

        $reader = $this->saveCheckRead($excel, $testFileName);
        $cells = $reader->readCells();

        $this->assertEquals($limit, mb_strlen($cells['A1']));
        // the limit counts characters, not bytes
        $this->assertEquals($limit, mb_strlen($cells['B1']));
        $this->assertEquals(str_repeat('я', $limit), $cells['B1']);
        // a value of exactly the limit is not touched, and the next cells are intact
        $this->assertEquals($limit, mb_strlen($cells['C1']));
        $this->assertEquals('short', $cells['D1']);
    }


    /**
     * Above the limit of hyperlinks Excel cannot read the relationships of the sheet,
     * so writing must stop with a clear error instead of producing a broken file
     */
    public function testHyperlinkLimit()
    {
        $excel = Excel::create(['Links']);
        $sheet = $excel->sheet();

        $exception = null;
        try {
            for ($i = 1; $i <= Sheet::MAX_HYPERLINKS + 5; $i++) {
                $sheet->writeRow(['link'], ['hyperlink' => 'https://example.com/' . $i]);
            }
        }
        catch (Exception $e) {
            $exception = $e;
        }

        $this->assertNotNull($exception, 'Too many hyperlinks must be rejected');
        $this->assertStringContainsString('Too many hyperlinks', $exception->getMessage());
        $this->assertCount(Sheet::MAX_HYPERLINKS, $sheet->getHyperlinks());
    }


    /**
     * Text and formulas of conditional formatting are escaped - "<" of an expression
     * or "&" of a searched text used to break the file
     */
    public function testConditionalFormattingEscaping()
    {
        $testFileName = __DIR__ . '/compat_cond_esc.xlsx';

        $excel = Excel::create(['Cond']);
        $sheet = $excel->sheet();
        $sheet->writeRow(['R&D <test>', 1]);
        $sheet->addConditionalFormatting('A1:A10', Conditional::contains('R&D <test>', ['fill-color' => '#FF0000']));
        $sheet->addConditionalFormatting('B1:B10', Conditional::expression('=B1<5', ['fill-color' => '#00FF00']));

        // the file must be readable, that is the whole point
        $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');

        $this->assertStringContainsString('text="R&amp;D &lt;test&gt;"', $xml);
        $this->assertStringContainsString('SEARCH(&quot;R&amp;D &lt;test&gt;&quot;,A1)', $xml);
        $this->assertStringContainsString('<formula>=B1&lt;5</formula>', $xml);
    }


    /**
     * The active cell of a <selection> must belong to the pane of that selection,
     * and <pane activePane="..."> must point to the pane holding the active cell
     */
    public function testFreezePanesSelection()
    {
        $testFileName = __DIR__ . '/compat_panes.xlsx';

        $cases = [
            // freeze, active cell, expected activePane, expected selection of that pane
            ['C5', null, 'bottomRight', 'C5'],
            ['C5', 'F20', 'bottomRight', 'F20'],
            ['C5', 'D2', 'topRight', 'D2'],
            ['C5', 'B7', 'bottomLeft', 'B7'],
            ['C5', 'A1', 'topLeft', 'A1'],
            ['A5', 'B2', 'topLeft', 'B2'],
            ['A5', null, 'bottomLeft', 'A5'],
            ['C1', 'A3', 'topLeft', 'A3'],
            ['C1', null, 'topRight', 'C1'],
        ];

        foreach ($cases as [$freeze, $activeCell, $expectedPane, $expectedCell]) {
            $excel = Excel::create(['Panes']);
            $sheet = $excel->sheet();
            $sheet->setFreeze($freeze);
            if ($activeCell) {
                $sheet->setActiveCell($activeCell);
            }
            $sheet->writeRow([1, 2, 3, 4]);

            $this->saveCheckRead($excel, $testFileName);
            $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');
            $message = 'freeze ' . $freeze . ', active ' . ($activeCell ?: 'default');

            $this->assertStringContainsString('activePane="' . $expectedPane . '"', $xml, $message);
            $this->assertStringContainsString('<selection pane="' . $expectedPane . '" activeCell="' . $expectedCell . '" sqref="' . $expectedCell . '"/>', $xml, $message);

            // every selection must keep its active cell inside its own sqref
            preg_match_all('#<selection pane="(\w+)" activeCell="([A-Z]+\d+)" sqref="([^"]+)"#', $xml, $m, PREG_SET_ORDER);
            $this->assertNotEmpty($m, $message);
            foreach ($m as $selection) {
                $firstCell = strpos($selection[3], ':') ? strstr($selection[3], ':', true) : $selection[3];
                $this->assertEquals($firstCell, $selection[2], $message . ': activeCell must be inside sqref');
            }

            unlink($testFileName);
        }

        // a range as the active cell keeps its top left cell active
        $excel = Excel::create(['Panes']);
        $sheet = $excel->sheet();
        $sheet->setFreeze('C5');
        $sheet->setActiveCell('D6:F9');
        $sheet->writeRow([1, 2, 3, 4]);

        $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');
        $this->assertStringContainsString('<selection pane="bottomRight" activeCell="D6" sqref="D6:F9"/>', $xml);
    }


    /**
     * A text format keeps the value as it is, without a quotePrefix
     */
    public function testTextFormat()
    {
        $testFileName = __DIR__ . '/compat_text.xlsx';

        $excel = Excel::create(['Text'], ['auto_convert_number' => true]);
        $sheet = $excel->sheet();
        $sheet->writeRow(['007', '1.50', '12345678901234567890'], ['format' => '@text']);

        $reader = $this->saveCheckRead($excel, $testFileName);
        $xml = $this->readXml($testFileName, 'xl/worksheets/sheet1.xml');
        $cells = $reader->readCells();

        $this->assertSame('007', $cells['A1']);
        $this->assertSame('1.50', $cells['B1']);
        $this->assertSame('12345678901234567890', $cells['C1']);
        // written as a string with the "@" number format, no quotePrefix attribute
        $this->assertStringNotContainsString('quotePrefix', $xml);
        $this->assertStringNotContainsString('quotePrefix', $this->readXml($testFileName, 'xl/styles.xml'));
        $this->assertStringContainsString('formatCode="@"', $this->readXml($testFileName, 'xl/styles.xml'));
    }
}
