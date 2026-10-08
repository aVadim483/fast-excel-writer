<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\Charts\Chart;
use avadim\FastExcelWriter\Charts\DataSeriesValues;
use avadim\FastExcelWriter\Charts\Properties;
use avadim\FastExcelWriter\Conditional\Conditional;
use avadim\FastExcelWriter\RichText\RichText;
use avadim\FastExcelWriter\RichText\RichTextFragment;
use avadim\FastExcelWriter\Style\Style;
use avadim\FastExcelWriter\Style\StyleManager;
use PHPUnit\Framework\TestCase;

final class ColorNoneTest extends TestCase
{
    private function parts(Excel $excel): array
    {
        $file = tempnam(sys_get_temp_dir(), 'color-none-');
        $zip = new ZipArchive();
        try {
            $excel->save($file);
            $this->assertTrue($zip->open($file));
            $parts = [];
            for ($index = 0; $index < $zip->numFiles; $index++) {
                $name = $zip->getNameIndex($index);
                if (substr($name, -4) === '.xml' || substr($name, -4) === '.vml') {
                    $parts[$name] = $zip->getFromIndex($index);
                    $this->assertStringNotContainsString('rgb=""', $parts[$name]);
                    $this->assertStringNotContainsString('rgb="none"', $parts[$name]);
                    $this->assertStringNotContainsString('<a:srgbClr val="none"', $parts[$name]);
                }
            }
            return $parts;
        }
        finally {
            $zip->close();
            unlink($file);
        }
    }

    private function cellStyle(array $parts, string $cell): SimpleXMLElement
    {
        $sheet = simplexml_load_string($parts['xl/worksheets/sheet1.xml']);
        $sheet->registerXPathNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $nodes = $sheet->xpath('//s:c[@r="' . $cell . '"]');
        $this->assertCount(1, $nodes, $cell . ': ' . $parts['xl/worksheets/sheet1.xml']);
        $styles = simplexml_load_string($parts['xl/styles.xml']);
        return $styles->cellXfs->xf[(int)$nodes[0]['s']];
    }

    public function testFillResetAcrossPublicApis(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setDefaultStyle(['fill-color' => '#ff0000']);
        $sheet->setColStyle('A', ['fill-color' => 'none']);
        $sheet->setRowStyle(2, ['bg-color' => 'none']);
        $sheet->writeTo('B1', 1)->applyFillGradient('#f00', '#00f')->applyFillColor('none');
        $sheet->writeTo('C1', 1)->applyBgColor(' NONE ');
        $sheet->writeTo('D1', 1)->applyStyle((new Style())->setFillColor('none'));
        $sheet->writeTo('E1', 1)->applyStyle((new Style())->setBgColor('none'));
        $sheet->writeTo('F1', 1)->applyFillGradient('none', '#00f');
        $sheet->writeTo('G1', 1)->applyStyle((new Style())->setFillGradient('#f00', 'none'));
        $sheet->setBgColor('H1', 'none');
        $sheet->setValue('A1', 1);
        $sheet->setValue('A2', 1);
        $sheet->setValue('B2', 1);
        $area = $sheet->beginArea('A3');
        $area->writeTo('A3', 1)->setBgColor('A3', 'none');
        $area->writeTo('B3', 1)->setBackgroundColor('B3', 'none');
        $area->writeTo('C3', 1)->applyBgColor('none');
        $area->writeTo('D3', 1)->applyStyle(['fill' => 'none']);
        $sheet->endAreas();
        $parts = $this->parts($excel);
        foreach (['A1','B1','C1','D1','E1','F1','G1','H1','A2','B2','A3','B3','C3','D3'] as $cell) {
            $this->assertSame(0, (int)$this->cellStyle($parts, $cell)['fillId'], $cell);
        }
    }

    public function testFontBorderTabAndRichTextReset(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->setDefaultFontColor('#ff0000');
        $sheet->setTabColor('#ff0000')->setTabColor('none');
        $sheet->writeTo('A1', 1)->applyFontColor('#00f')->applyTextColor('none')->applyBorder('thin', 'none');
        $sheet->writeTo('B1', 1)->applyStyle((new Style())->setColor('none')->setBorderLeft('double', 'none'));
        $area = $sheet->beginArea('C1');
        $area->writeTo('C1', 1)->setFgColor('C1', 'none');
        $sheet->endAreas();
        $fragment = (new RichTextFragment('text'))->setColor('#00f')->setColor('none');
        $this->assertStringContainsString('<color auto="1"/>', $fragment->outXml());
        $rich = new RichText('<c="none">text</c>');
        $this->assertStringContainsString('<color auto="1"/>', $rich->outXml());
        $parts = $this->parts($excel);
        $styles = simplexml_load_string($parts['xl/styles.xml']);
        foreach (['A1','B1','C1'] as $cell) {
            $xf = $this->cellStyle($parts, $cell);
            $this->assertSame('1', (string)$styles->fonts->font[(int)$xf['fontId']]->color['auto'], $cell);
        }
        $xf = $this->cellStyle($parts, 'A1');
        $border = $styles->borders->border[(int)$xf['borderId']];
        $this->assertSame('thin', (string)$border->left['style']);
        $this->assertSame('1', (string)$border->left->color['auto']);
        $this->assertStringNotContainsString('<tabColor', $parts['xl/worksheets/sheet1.xml']);
    }

    public function testConditionalAndNoteColors(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->writeRow([1, 2, 3]);
        $sheet->addConditionalFormatting('A1:A3', Conditional::greaterThan(0)->setFillColor('none')->setFontColor('none'));
        $sheet->addConditionalFormatting('B1:B3', Conditional::colorScale('none', '#00f'));
        $sheet->addConditionalFormatting('C1:C3', Conditional::dataBar('none'));
        $sheet->addNote('A1', 'note', ['fill_color' => 'none']);
        $sheet->addNote('B1', 'note', ['bg_color' => 'none']);
        $parts = $this->parts($excel);
        $this->assertStringContainsString('<font><color auto="1"/></font><fill><patternFill patternType="none"/></fill>', $parts['xl/styles.xml']);
        $this->assertSame(2, substr_count($parts['xl/worksheets/sheet1.xml'], '<color auto="1"/>'));
        $vml = implode('', array_filter($parts, static function ($name) { return substr($name, -4) === '.vml'; }, ARRAY_FILTER_USE_KEY));
        $this->assertSame(2, substr_count($vml, 'filled="f"'));
        $this->assertStringNotContainsString('fillcolor="none"', $vml);
    }

    public function testChartColors(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->writeRow(['Name', 'Value']);
        $sheet->writeRow(['One', 1]);
        $sheet->writeRow(['Two', 2]);
        $chart = Chart::make(Chart::TYPE_LINE, 'Line', ['B1' => 'B2:B3'])->setCategoryAxisLabels('A2:A3');
        $series = $chart->getPlotArea()->getPlotDataSeriesByIndex(0)->getDataSeriesValues()[0];
        $series->setColor('none');
        $axis = $chart->getChartAxisX();
        $axis->setFillParameters('none')->setLineParameters('none');
        $axis->setGlowProperties(2, 'none');
        $axis->setShadowProperties(Properties::SHADOW_PRESETS_OUTER_BOTTOM, 'none');
        $grid = $chart->getMajorGridlines();
        $grid->setLineColorProperties('none');
        $grid->setGlowProperties(2, 'none');
        $grid->setShadowProperties(Properties::SHADOW_PRESETS_OUTER_BOTTOM, 'none');
        $chart = (new Chart('Line', $chart->getPlotArea(), null, true, '0', null, null, $axis, null, $grid))
            ->setChartType(Chart::TYPE_LINE)->setCategoryAxisLabels('A2:A3');
        $sheet->addChart('D1:L12', $chart);
        $pie = Chart::make(Chart::TYPE_PIE, 'Pie',
            new DataSeriesValues('B2:B3', null, ['segment_colors' => ['none', '#f00']]));
        $sheet->addChart('D14:L26', $pie);
        $parts = $this->parts($excel);
        $this->assertStringContainsString('<a:noFill/>', $parts['xl/charts/chart1.xml']);
        $this->assertStringContainsString('<a:alpha val="0"/>', $parts['xl/charts/chart1.xml']);
        $this->assertStringContainsString('<a:outerShdw', $parts['xl/charts/chart1.xml']);
        $this->assertStringContainsString('<a:noFill/>', $parts['xl/charts/chart2.xml']);
        $this->assertStringContainsString('val="ff0000"', $parts['xl/charts/chart2.xml']);
    }

    public function testInvalidColorsStillFail(): void
    {
        $this->assertNull(StyleManager::normalizeColor('none'));
        $this->assertNull(StyleManager::normalizeColor(' NONE '));
        $this->expectException(\avadim\FastExcelWriter\Exceptions\Exception::class);
        StyleManager::normalizeColor('not-a-color');
    }

    public function testColorsCanBeReappliedAfterNone(): void
    {
        $excel = Excel::create();
        $sheet = $excel->sheet();
        $sheet->writeTo('A1', 1)->applyFillColor('none')->applyFillGradient('#f00', '#00f');
        $sheet->writeTo('B1', 1)->applyFillGradient('none', '#00f')->applyFillColor('#0f0');
        $sheet->writeTo('C1', 1)->applyFontColor('none')->applyFontColor('#00f');
        $parts = $this->parts($excel);
        $styles = simplexml_load_string($parts['xl/styles.xml']);
        $xf = $this->cellStyle($parts, 'A1');
        $this->assertCount(2, $styles->fills->fill[(int)$xf['fillId']]->gradientFill->stop);
        $xf = $this->cellStyle($parts, 'B1');
        $this->assertSame('FF00FF00', (string)$styles->fills->fill[(int)$xf['fillId']]->patternFill->fgColor['rgb']);
        $xf = $this->cellStyle($parts, 'C1');
        $this->assertSame('FF0000FF', (string)$styles->fonts->font[(int)$xf['fontId']]->color['rgb']);
    }
}
