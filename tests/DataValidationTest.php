<?php

declare(strict_types=1);

namespace avadim\FastExcelWriter\Tests;

use avadim\FastExcelWriter\DataValidation\DataValidation;
use avadim\FastExcelWriter\Conditional\Conditional;
use avadim\FastExcelWriter\Excel;
use PHPUnit\Framework\TestCase;

final class DataValidationTest extends TestCase
{
    private function worksheetsXml(string $file): string
    {
        $zip = new \ZipArchive();
        $zip->open($file);
        $xml = '';
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = $zip->getNameIndex($i);
            if (preg_match('#worksheets/sheet\d+\.xml$#', $name)) {
                $xml .= $zip->getFromName($name);
            }
        }
        $zip->close();

        return $xml;
    }

    public function testReuseSameValidationForMultipleRanges()
    {
        // issue #137: reusing one DataValidation object for several ranges must keep every range
        $excel = Excel::create(['test']);
        $sheet = $excel->sheet();
        $validation = DataValidation::integer('>=', 0);
        $sheet->addDataValidation('A1:A100', $validation);
        $sheet->addDataValidation('C1:C100', $validation);

        $file = __DIR__ . '/tmp/issue137_dv.xlsx';
        $excel->save($file);
        $xml = $this->worksheetsXml($file);
        unlink($file);

        $this->assertStringContainsString('sqref="A1:A100"', $xml);
        $this->assertStringContainsString('sqref="C1:C100"', $xml);
    }

    public function testReuseSameConditionalForMultipleRanges()
    {
        // issue #137: the same aliasing problem applies to reused Conditional objects
        $excel = Excel::create(['test']);
        $sheet = $excel->sheet();
        $conditional = Conditional::greaterThan(10);
        $sheet->addConditionalFormatting('A1:A10', $conditional);
        $sheet->addConditionalFormatting('C1:C10', $conditional);

        $file = __DIR__ . '/tmp/issue137_cf.xlsx';
        $excel->save($file);
        $xml = $this->worksheetsXml($file);
        unlink($file);

        $this->assertStringContainsString('sqref="A1:A10"', $xml);
        $this->assertStringContainsString('sqref="C1:C10"', $xml);
    }

    public function testInteger()
    {
        $validation = DataValidation::integer('between', [1, 10]);
        $attributes = $validation->getAttributes();
        
        $this->assertEquals('whole', $attributes['type']);
        $this->assertEquals('between', $attributes['operator']);
        
        $xml = $validation->toXml();
        $this->assertStringContainsString('type="whole"', $xml);
        $this->assertStringContainsString('operator="between"', $xml);
        $this->assertStringContainsString('<formula1>1</formula1>', $xml);
        $this->assertStringContainsString('<formula2>10</formula2>', $xml);
    }

    public function testList()
    {
        // List from array
        $validation = DataValidation::list(['item1', 'item2', 'item3']);
        $attributes = $validation->getAttributes();
        
        $this->assertEquals('list', $attributes['type']);
        
        $xml = $validation->toXml();
        $this->assertStringContainsString('type="list"', $xml);
        $this->assertStringContainsString('<formula1>&quot;item1,item2,item3&quot;</formula1>', $xml);

        // List from string (range)
        $validation = DataValidation::list('=$A$1:$A$10');
        $xml = $validation->toXml();
        $this->assertStringContainsString('<formula1>$A$1:$A$10</formula1>', $xml);
    }

    public function testCustom()
    {
        $validation = DataValidation::custom('=A1>0');
        $attributes = $validation->getAttributes();
        
        $this->assertEquals('custom', $attributes['type']);
        
        $xml = $validation->toXml();
        $this->assertStringContainsString('type="custom"', $xml);
        $this->assertStringContainsString('<formula1>A1&gt;0</formula1>', $xml);
    }

    public function testAttributes()
    {
        $validation = DataValidation::integer('>', 0)
            ->allowBlank(true)
            ->setError('Error Message', 'Error Title')
            ->setPrompt('Prompt Message', 'Prompt Title')
            ->setErrorStyle(DataValidation::STYLE_WARNING);
        
        $attributes = $validation->getAttributes();
        
        $this->assertEquals(1, $attributes['allowBlank']);
        $this->assertEquals('Error Message', $attributes['error']);
        $this->assertEquals('Error Title', $attributes['errorTitle']);
        $this->assertEquals('Prompt Message', $attributes['prompt']);
        $this->assertEquals('Prompt Title', $attributes['promptTitle']);
        $this->assertEquals('warning', $attributes['errorStyle']);
        $this->assertEquals(1, $attributes['showErrorMessage']);
        $this->assertEquals(1, $attributes['showInputMessage']);
    }

    public function testDecimal()
    {
        $validation = DataValidation::decimal('notBetween', [0.1, 0.9]);
        $xml = $validation->toXml();
        
        $this->assertStringContainsString('type="decimal"', $xml);
        $this->assertStringContainsString('operator="notBetween"', $xml);
        $this->assertStringContainsString('<formula1>0.1</formula1>', $xml);
        $this->assertStringContainsString('<formula2>0.9</formula2>', $xml);
    }

    public function testDate()
    {
        $validation = DataValidation::date('equal', '2023-01-01');
        $xml = $validation->toXml();
        
        $this->assertStringContainsString('type="date"', $xml);
        $this->assertStringContainsString('operator="equal"', $xml);
        $this->assertStringContainsString('<formula1>2023-01-01</formula1>', $xml);
    }

    public function testTextLength()
    {
        $validation = DataValidation::textLength('lessThanOrEqual', 20);
        $xml = $validation->toXml();
        
        $this->assertStringContainsString('type="textLength"', $xml);
        $this->assertStringContainsString('operator="lessThanOrEqual"', $xml);
        $this->assertStringContainsString('<formula1>20</formula1>', $xml);
    }

    public function testTime()
    {
        $validation = DataValidation::make(DataValidation::TYPE_TIME)
            ->setOperator('between', '09:00', '18:00');
        $xml = $validation->toXml();

        $this->assertStringContainsString('type="time"', $xml);
        $this->assertStringContainsString('operator="between"', $xml);
        $this->assertStringContainsString('<formula1>09:00</formula1>', $xml);
        $this->assertStringContainsString('<formula2>18:00</formula2>', $xml);
    }

    public function testInvalidOperator()
    {
        $this->expectException(\avadim\FastExcelWriter\Exceptions\ExceptionDataValidation::class);
        DataValidation::integer('invalid_operator', 10);
    }

    public function testInvalidStyle()
    {
        $validation = DataValidation::integer('>', 0);
        $this->expectException(\avadim\FastExcelWriter\Exceptions\ExceptionDataValidation::class);
        $validation->setErrorStyle('invalid_style');
    }
}
