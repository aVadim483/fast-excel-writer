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
                    $properties[$property->nodeName] = $property->getAttribute('val')
                        ?: $property->getAttribute('rgb');
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
}
