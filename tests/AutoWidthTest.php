<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Style\Font;
use avadim\FastExcelWriter\Style\WidthCache;
use PHPUnit\Framework\TestCase;

final class AutoWidthTest extends TestCase
{
    protected function tearDown(): void
    {
        Font::setFontWidthFactor(1.0);
    }

    public function testFactorChangesAffectCachedWidths(): void
    {
        $width = Font::calcTextWidth('Calibri', 11, 'cached text');
        Font::setFontWidthFactor(1.5);
        $this->assertEqualsWithDelta($width * 1.5, Font::calcTextWidth('Calibri', 11, 'cached text'), 0.000001);
    }

    public function testCacheEvictionAndKeyBudgets(): void
    {
        $cache = new WidthCache(2, 6, 4);
        $cache->put('aa', 0);
        $cache->put('bb', 2);
        $this->assertSame(0, $cache->get('aa'));
        $cache->put('cccc', 3);
        $this->assertNull($cache->get('aa'));
        $this->assertSame(2, $cache->get('bb'));
        $cache->put('long-key', 4);
        $this->assertNull($cache->get('long-key'));
        for ($i = 0; $i < 20; ++$i) {
            $cache->put((string)$i, $i);
        }
        $this->assertSame(19, $cache->get('19'));
        $this->assertNull($cache->get('0'));
    }

    public function testUniqueValuesDoNotKeepGrowingMemory(): void
    {
        $font = new Font();
        for ($i = 0; $i < 10000; ++$i) {
            $font->calculateTextWidth('Calibri', 11, (string)$i);
        }
        $memory = memory_get_usage();
        for ($i = 10000; $i < 60000; ++$i) {
            $font->calculateTextWidth('Calibri', 11, (string)$i);
        }
        $this->assertLessThan(1048576, memory_get_usage() - $memory);
    }

    public function testWorkbookCachesAreIndependent(): void
    {
        $first = Excel::create();
        $second = Excel::create();
        $this->assertNotSame($first->getAutoWidthFont(), $second->getAutoWidthFont());
        $this->assertSame($first->getAutoWidthFont(), $first->getAutoWidthFont());
    }
}
