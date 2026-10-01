<?php

declare(strict_types=1);

use avadim\FastExcelWriter\Excel;
use avadim\FastExcelWriter\RichText\RichText;
use avadim\FastExcelReader\Excel as ExcelReader;
use PHPUnit\Framework\TestCase;

final class RichTextTest extends TestCase
{
    private function runs(RichText $text): array
    {
        $doc = new DOMDocument();
        $this->assertTrue($doc->loadXML('<root>' . $text->outXml() . '</root>'));
        $runs = [];
        foreach ($doc->getElementsByTagName('r') as $run) {
            $properties = [];
            foreach ($run->getElementsByTagName('rPr') as $rPr) {
                foreach ($rPr->childNodes as $property) {
                    $properties[$property->nodeName] = $property->hasAttribute('val')
                        ? $property->getAttribute('val') : $property->getAttribute('rgb');
                }
            }
            $runs[] = [$run->textContent, $properties];
        }
        return $runs;
    }

    public function testFluentScriptsAndMutations(): void
    {
        $text = new RichText();
        $text->addText('H')->addText('2')->setSubscript()->setBold()->setColor('red');
        $text->addText('O');
        $this->assertSame([
            ['H', []],
            ['2', ['b' => '', 'color' => 'FFFF0000', 'vertAlign' => 'subscript']],
            ['O', []],
        ], $this->runs($text));
        $text->fragment(1)->setSuperscript();
        $this->assertSame('superscript', $this->runs($text)[1][1]['vertAlign']);
        $text->fragment(1)->setBaseline();
        $this->assertSame('baseline', $this->runs($text)[1][1]['vertAlign']);
        $text->addText('3')->setSuperscript();
        $this->assertSame('superscript', $this->runs($text)[3][1]['vertAlign']);
        $text->setSubscript()->setBaseline();
        $this->assertSame('baseline', $this->runs($text)[3][1]['vertAlign']);
    }

    public function testTagsAndNestedScripts(): void
    {
        $text = new RichText('H<sub>2</sub>O m<sup><b>2</b></sup>!');
        $this->assertSame([
            ['H', []], ['2', ['vertAlign' => 'subscript']], ['O m', []],
            ['2', ['b' => '', 'vertAlign' => 'superscript']], ['!', []],
        ], $this->runs($text));
        $nested = new RichText('<sup>a<sub>b</sub>c<sup>d</sup>e</sup>f');
        $this->assertSame([
            ['a', ['vertAlign' => 'superscript']], ['b', ['vertAlign' => 'subscript']],
            ['c', ['vertAlign' => 'superscript']], ['d', ['vertAlign' => 'superscript']],
            ['e', ['vertAlign' => 'superscript']], ['f', []],
        ], $this->runs($nested));
        $split = new RichText('<sub>0');
        $split->outXml();
        $split->addTaggedText('</sub>1');
        $this->assertSame([['0', ['vertAlign' => 'subscript']], ['1', []]], $this->runs($split));
    }

    public function testExistingTagAliases(): void
    {
        $short = new RichText('<s=18><f=Arial><c=red><b><i><u>x<sup>0</sup>y</u></i></b></c></f></s>z');
        $long = new RichText('<size=18><font=Arial><color=red><bold><italic><underline>x<sup>0</sup>y</underline></italic></bold></color></font></size>z');
        $this->assertSame($short->outXml(), $long->outXml());
        $runs = $this->runs($short);
        $this->assertSame('18', $runs[1][1]['sz']);
        $this->assertSame('single', $runs[1][1]['u']);
        $this->assertSame('superscript', $runs[1][1]['vertAlign']);
        $this->assertArrayNotHasKey('vertAlign', $runs[2][1]);
        $this->assertSame(['z', []], $runs[3]);
    }

    public function testScriptsInSavedWorkbook(): void
    {
        foreach ([true, false] as $sharedStrings) {
            $path = tempnam(sys_get_temp_dir(), 'rich-text-143-');
            try {
                $excel = Excel::create(['Scripts'], ['shared_string' => $sharedStrings]);
                $excel->sheet()->writeRow([new RichText('H<sub>2</sub>O'), new RichText('m<sup>2</sup>')]);
                $excel->save($path);
                $this->assertTrue(ExcelReader::validate($path, $errors));
                $zip = new ZipArchive();
                $this->assertTrue($zip->open($path));
                $xml = $zip->getFromName('xl/sharedStrings.xml');
                $zip->close();
                $doc = new DOMDocument();
                $this->assertTrue($doc->loadXML($xml));
                $alignments = $doc->getElementsByTagName('vertAlign');
                $this->assertSame(2, $alignments->length);
                $this->assertSame('subscript', $alignments->item(0)->getAttribute('val'));
                $this->assertSame('superscript', $alignments->item(1)->getAttribute('val'));
                $strings = $doc->getElementsByTagName('si');
                $this->assertSame('H2O', $strings->item(0)->textContent);
                $this->assertSame('m2', $strings->item(1)->textContent);
            }
            finally {
                unlink($path);
            }
        }
    }

    public function testStrikethroughAndUnderline(): void
    {
        $text = new RichText();
        $text->addText('old')->setStrike()->setUnderline();
        $text->addText('new')->setUnderline(true);
        $this->assertSame([
            ['old', ['u' => 'single', 'strike' => '']],
            ['new', ['u' => 'double']],
        ], $this->runs($text));
        $text->fragment(1)->setUnderline(false);
        $this->assertSame('single', $this->runs($text)[1][1]['u']);
        $this->assertSame([
            ['a', ['strike' => '']], ['b', ['strike' => '']],
            ['c', ['strike' => '']], ['d', []], ['18', ['sz' => '18']],
        ], $this->runs(new RichText('<strike>a<del>b</del>c</strike>d<s=18>18</s>')));
    }

    public function testLiteralAndTaggedTextEscaping(): void
    {
        $plain = (new RichText())->addText('A & B, x < y, <b>literal</b>, &amp; "quoted"');
        $this->assertSame('A & B, x < y, <b>literal</b>, &amp; "quoted"', $this->runs($plain)[0][0]);
        $tagged = new RichText('<b>A &amp; B, x &lt; y, &lt;b&gt;, &amp;amp; &#34;quoted&#34;</b>');
        $this->assertSame([['A & B, x < y, <b>, &amp; "quoted"', ['b' => '']]], $this->runs($tagged));
        $this->assertSame('x < y & z <unknown>value</unknown>', $this->runs(new RichText('x < y & z <unknown>value</unknown>'))[0][0]);

        $font = 'Example "Sans" & <Alt>';
        $plain->setFont($font);
        $this->assertSame($font, $this->runs($plain)[0][1]['rFont']);
        $tagged = new RichText('<font="Example &quot;Sans&quot; &amp; &lt;Alt&gt;">x</font>');
        $this->assertSame($font, $this->runs($tagged)[0][1]['rFont']);
        $this->assertSame('A > B', $this->runs(new RichText('<font="A > B">x</font>'))[0][1]['rFont']);
    }

    public function testExplicitFormattingAndFractionalSize(): void
    {
        $text = (new RichText())->addText('plain')->setBold(false)->setItalic(false)
            ->setStrike(false)->removeUnderline()->setSize(10.5);
        $this->assertSame([['plain', [
            'b' => '0', 'i' => '0', 'u' => 'none', 'strike' => '0', 'sz' => '10.5',
        ]]], $this->runs($text));
        $text->fragment(0)->setBold(true)->setItalic(true)->setStrike(true)->setUnderline(true)->setSize(12);
        $this->assertSame([['plain', [
            'b' => '', 'i' => '', 'u' => 'double', 'strike' => '', 'sz' => '12',
        ]]], $this->runs($text));
        $this->assertSame('10.5', $this->runs(new RichText('<s=10.5>x</s>'))[0][1]['sz']);
    }

    public function testInvalidFontSizesAreRejected(): void
    {
        foreach ([0, -1, INF, NAN] as $size) {
            try {
                (new RichText())->addText('x')->setSize($size);
                $this->fail('Invalid size must be rejected');
            }
            catch (InvalidArgumentException $e) {
                $this->assertStringContainsString('Font size', $e->getMessage());
            }
        }
        $this->expectException(InvalidArgumentException::class);
        new RichText('<s=invalid>x</s>');
    }

    public function testEscapingInWorkbookAndNotes(): void
    {
        $path = tempnam(sys_get_temp_dir(), 'rich-escaping-');
        try {
            $value = "A & <B>\r\x07 _x000D_ &amp;";
            $rich = (new RichText())->addText($value)->setFont('Example "Sans" & Co')
                ->setSize(10.5)->setBold(false)->setItalic(false)->setStrike(false)->removeUnderline();
            $excel = Excel::create();
            $sheet = $excel->sheet();
            $sheet->writeRow([$rich, new RichText('<b>A &amp; B</b>')]);
            $sheet->addNote('A1', $value);
            $sheet->addNote('B1', $rich);
            $excel->save($path);
            $this->assertTrue(ExcelReader::validate($path, $errors));
            $cells = ExcelReader::open($path)->readCells();
            $this->assertSame($value, $cells['A1']);
            $this->assertSame('A & B', $cells['B1']);
            $zip = new ZipArchive();
            $zip->open($path);
            foreach (['xl/sharedStrings.xml', 'xl/comments1.xml'] as $entry) {
                $xml = $zip->getFromName($entry);
                $doc = new DOMDocument();
                $this->assertTrue($doc->loadXML($xml));
                $this->assertSame('Example "Sans" & Co', $doc->getElementsByTagName('rFont')->item(0)->getAttribute('val'));
                $this->assertSame('10.5', $doc->getElementsByTagName('sz')->item(0)->getAttribute('val'));
                $this->assertStringContainsString('_x000D_', $xml);
                $this->assertStringContainsString('_x0007_', $xml);
                $this->assertStringContainsString('_x005F_x000D_', $xml);
                $this->assertStringContainsString('<b val="0"/>', $xml);
            }
            $zip->close();
        }
        finally {
            unlink($path);
        }
    }
}
