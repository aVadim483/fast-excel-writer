<?php

namespace avadim\FastExcelWriter\RichText;

use avadim\FastExcelWriter\Style\StyleManager;

class RichText
{
    protected string $text = '';
    protected array $buffer;
    protected int $pos;
    protected int $cnt = -1;
    protected array $prop = ['b' => null, 'i' => null, 'u' => null, 'f' => null, 's' => null, 'c' => null];
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
        if ($text) {
            $this->buffer = mb_str_split($text);
            $this->pos = 0;
            while (isset($this->buffer[$this->pos])) {
                $token = $this->getToken();
                if ($token) {
                    if (substr($token, 0, 2) === '</') {
                        $tag = substr($token, 2, 1);
                        if (isset($this->prop[$tag])) {
                            $this->prop[$tag] = null;
                        }
                    }
                    elseif ($token[0] === '<') {
                        switch (substr($token, 0, 2)) {
                            case '<b':
                                $this->prop['b'] = true;
                                break;
                            case '<i':
                                $this->prop['i'] = true;
                                break;
                            case '<u':
                                $this->prop['u'] = true;
                                break;
                            case '<f':
                                if (strpos($token, '=')) {
                                    [$name, $arg] = explode('=', $token, 2);
                                    $this->prop['f'] = trim($arg, '"\'>');
                                }
                                break;
                            case '<s':
                                if (strpos($token, '=')) {
                                    [$name, $arg] = explode('=', $token, 2);
                                    $this->prop['s'] = trim($arg, '"\'>');
                                }
                                break;
                            case '<c':
                                if (strpos($token, '=')) {
                                    [$name, $arg] = explode('=', $token, 2);
                                    $this->prop['c'] = StyleManager::normalizeColor(trim($arg, '"\'>'));
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
     * Add tagged text (<b>, <i>, <u>, <f>, <s>, <c>)
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
        if (!$this->xml) {
            $this->xml = '';
            foreach ($this->fragments as $fragment) {
                $this->xml .= $fragment->outXml();
            }
        }

        return $this->xml;
    }
}