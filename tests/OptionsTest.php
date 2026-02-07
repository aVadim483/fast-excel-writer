<?php

namespace avadim\FastExcelWriter;

use PHPUnit\Framework\TestCase;

class OptionsTest extends TestCase
{
    public function testConstructorWithArray()
    {
        $options = [
            'temp_dir' => __DIR__ . '/tmp',
            'shared_string' => true,
        ];
        $excel = new Excel($options);
        $this->assertInstanceOf(Excel::class, $excel);
        
        $writer = $excel->getWriter();
        // Внутренние свойства писателя не всегда доступны напрямую, 
        // но мы можем проверить, что объект создался без ошибок.
        $this->assertNotNull($writer);
    }

    public function testConstructorWithOptionsObject()
    {
        $options = Options::create()
            ->tempDir(__DIR__ . '/tmp')
            ->sharedString(true);
            
        $excel = new Excel($options);
        $this->assertInstanceOf(Excel::class, $excel);
        
        $writer = $excel->getWriter();
        $this->assertNotNull($writer);
    }

    public function testStaticCreateWithArray()
    {
        $options = [
            'temp_dir' => __DIR__ . '/tmp',
        ];
        $excel = Excel::create('Sheet1', $options);
        $this->assertInstanceOf(Excel::class, $excel);
        $this->assertEquals('Sheet1', $excel->sheet()->getName());
    }

    public function testStaticCreateWithOptionsObject()
    {
        $options = Options::create()->tempDir(__DIR__ . '/tmp');
        $excel = Excel::create('Sheet1', $options);
        $this->assertInstanceOf(Excel::class, $excel);
        $this->assertEquals('Sheet1', $excel->sheet()->getName());
    }

    public function testOptionsArrayAccess()
    {
        $options = new Options(['a' => 1]);
        $this->assertEquals(1, $options['a']);
        $this->assertTrue(isset($options['a']));
        
        $options['b'] = 2;
        $this->assertEquals(2, $options['b']);
        
        unset($options['a']);
        $this->assertFalse(isset($options['a']));
    }

    public function testOptionsFluentInterface()
    {
        $options = Options::create()
            ->tempDir('/tmp')
            ->tempPrefix('pre_')
            ->autoConvertNumber(true)
            ->sharedString(true)
            ->locale('ru')
            ->defaultFont(['font-name' => 'Arial'])
            ->writerClass('MyWriter')
            ->styleManagerClass('MyStyleManager');

        $array = $options->toArray();
        $this->assertEquals('/tmp', $array['temp_dir']);
        $this->assertEquals('pre_', $array['temp_prefix']);
        $this->assertTrue($array['auto_convert_number']);
        $this->assertTrue($array['shared_string']);
        $this->assertEquals('ru', $array['locale']);
        $this->assertEquals(['font-name' => 'Arial'], $array['default_font']);
        $this->assertEquals('MyWriter', $array['writer_class']);
        $this->assertEquals('MyStyleManager', $array['style_manager']);
    }
}
