<?php

namespace avadim\FastExcelWriter\RichText;

use avadim\FastExcelWriter\Style\StyleManager;

class RichText
{
    protected string $text = '';
    protected array $buffer;
    protected int $pos;
    protected int $cnt = -1;
    protected array $prop = ['b' => null, 'i' => null, 'u' => null, 'f' => null, 'sz' => null, 'c' => null, 'vertAlign' => null, 'strike' => null];
    protected array $propStacks = [];
    protected array $fragments = [];
    protected ?string $xml = null;

    /**
     * RichText constructor
     *
     * @param string|array|null $fragments
     */
    public function __construct(...$fragments)
    {
        foreach ($fragments as $item) {
            $this->addTaggedText($item);
        }
    }


    /**
     * @return string
     */
    protected function getToken(): string
    {
        $token = '';
        if ((isset($this->buffer[$this->pos]) && $this->buffer[$this->pos] === '<')) {
            $tag = true;
            $breakChar = '>';
        }
        else {
            $tag = false;
            $breakChar = '<';
        }
        while (isset($this->buffer[$this->pos]) && $this->buffer[$this->pos] !== $breakChar) {
            $token .= $this->buffer[$this->pos++];
        }
        if ($tag && isset($this->buffer[$this->pos]) && $this->buffer[$this->pos] === '>') {
            $token .= $this->buffer[$this->pos++];
        }

        return $token;
    }

    /**
     * @param string $text
     *
     * @return array
     */
    protected function parse(string $text): array
    {
        $fragments = [];
        if ($text !== '') {
            $this->buffer = mb_str_split($text);
            $this->pos = 0;
            while (isset($this->buffer[$this->pos])) {
                $token = $this->getToken();
                if ($token !== '') {
                    if ($token[0] === '<') {
                        $tags = [
                            'b' => 'b', 'bold' => 'b', 'i' => 'i', 'italic' => 'i',
                            'u' => 'u', 'underline' => 'u', 'f' => 'f', 'font' => 'f',
                            's' => 'sz', 'size' => 'sz', 'c' => 'c', 'color' => 'c',
                            'sub' => 'vertAlign', 'sup' => 'vertAlign',
                            'strike' => 'strike', 'del' => 'strike',
                        ];
                        if (preg_match('~^<(/?)([a-z]+)(?:=(.*))?>$~i', $token, $match)) {
                            $tag = strtolower($match[2]);
                            if (!isset($tags[$tag])) {
                                continue;
                            }
                            $key = $tags[$tag];
                            if ($match[1] === '/') {
                                $this->prop[$key] = !empty($this->propStacks[$key])
                                    ? array_pop($this->propStacks[$key]) : null;
                            }
                            else {
                                if (in_array($key, ['f', 'sz', 'c'], true)) {
                                    if (!isset($match[3])) {
                                        continue;
                                    }
                                    $value = trim($match[3], '"\'');
                                    if ($key === 'c') {
                                        $value = StyleManager::normalizeColor($value);
                                    }
                                }
                                else {
                                    $value = $key === 'u' ? 'single' : true;
                                    if ($key === 'vertAlign') {
                                        $value = $tag === 'sub' ? 'subscript' : 'superscript';
                                    }
                                }
                                $this->propStacks[$key][] = $this->prop[$key];
                                $this->prop[$key] = $value;
                            }
                        }
                    }
                    else {
                        $fragments[] = new RichTextFragment($token, $this->prop);
                    }
                }
            }
            $this->xml = null;
        }

        return $fragments;
    }

    /**
     * Add a text fragment
     *
     * @param string $text
     * @param mixed $prop
     *
     * @return $this
     */
    public function addText(string $text, $prop = []): RichText
    {
        $fragment = new RichTextFragment($text, $prop);
        $this->fragments[++$this->cnt] = $fragment;

        return $this;
    }

    /**
     * Add tagged text (<b>, <i>, <u>, <f>, <s>, <c>, <sub>, <sup>, <strike>, <del>)
     *
     * @param string $text
     *
     * @return RichText
     */
    public function addTaggedText(string $text): RichText
    {
        $fragments = $this->parse($text);
        foreach ($fragments as $fragment) {
            $this->fragments[++$this->cnt] = $fragment;
        }

        return $this;
    }

    /**
     * Set bold font for the last added fragment
     *
     * @return $this
     */
    public function setBold(): RichText
    {
        $this->fragments[$this->cnt]->setBold();

        return $this;
    }

    /** Set strikethrough for the last added fragment. */
    public function setStrike(): RichText
    {
        $this->fragments[$this->cnt]->setStrike();

        return $this;
    }

    /** Set subscript for the last added fragment. */
    public function setSubscript(): RichText
    {
        $this->fragments[$this->cnt]->setSubscript();

        return $this;
    }

    /** Set superscript for the last added fragment. */
    public function setSuperscript(): RichText
    {
        $this->fragments[$this->cnt]->setSuperscript();

        return $this;
    }

    /** Restore the baseline for the last added fragment. */
    public function setBaseline(): RichText
    {
        $this->fragments[$this->cnt]->setBaseline();

        return $this;
    }

    /**
     * Set italic font for the last added fragment
     *
     * @return $this
     */
    public function setItalic(): RichText
    {
        $this->fragments[$this->cnt]->setItalic();

        return $this;
    }

    /**
     * Set underline for the last added fragment
     *
     * @param bool|null $double
     *
     * @return $this
     */
    public function setUnderline(?bool $double = false): RichText
    {
        $this->fragments[$this->cnt]->setUnderline($double);

        return $this;
    }

    /**
     * Set font name for the last added fragment
     *
     * @param string $font
     *
     * @return $this
     */
    public function setFont(string $font): RichText
    {
        $this->fragments[$this->cnt]->setFont($font);

        return $this;
    }

    /**
     * Set font size for the last added fragment
     *
     * @param int $size
     *
     * @return $this
     */
    public function setSize(int $size): RichText
    {
        $this->fragments[$this->cnt]->setSize($size);

        return $this;
    }

    /**
     * Set font color for the last added fragment
     *
     * @param string $color
     *
     * @return $this
     */
    public function setColor(string $color): RichText
    {
        $this->fragments[$this->cnt]->setColor($color);

        return $this;
    }

    /**
     * Get all fragments
     *
     * @return array
     */
    public function fragments(): array
    {
        return $this->fragments;
    }

    /**
     * Get fragment by its index
     *
     * @param $num
     *
     * @return RichTextFragment
     */
    public function fragment($num): RichTextFragment
    {
        return $this->fragments[$num];
    }

    /**
     * @return string
     */
    public function __toString()
    {
        return $this->outXml();
    }

    /**
     * Returns XML representation
     *
     * @return string
     */
    public function outXml(): string
    {
        // Fragments can be changed directly through fragment() or fragments().
        $this->xml = '';
        foreach ($this->fragments as $fragment) {
            $this->xml .= $fragment->outXml();
        }

        return $this->xml;
    }
}
