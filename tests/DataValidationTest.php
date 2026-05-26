<?php

declare(strict_types=1);

namespace avadim\FastExcelWriter\Tests;

use avadim\FastExcelWriter\DataValidation\DataValidation;
use PHPUnit\Framework\TestCase;

final class DataValidationTest extends TestCase
{
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
