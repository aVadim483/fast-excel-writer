<?php

namespace avadim\FastExcelWriter;

class RichText
{
    protected string $text = '';
    protected array $buffer;
    protected int $pos;
    protected array $prop = ['b' => null, 'i' => null, 'u' => null, 'f' => null, 's' => null, 'c' => null];
    protected array $fragments = [];
    protected ?string $xml = null;

    /**
     * @param string|null $text
     */
    public function __construct(?string $text = '')
    {
        $this->resetProps();
        $this->parse($text);
    }


    protected function resetProps()
    {
        $this->prop = ['b' => null, 'i' => null, 'u' => null, 'f' => null, 's' => null, 'c' => null];
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
                        $this->fragments[] = ['text' => $token, 'prop' => $this->prop];
                    }
                }
            }
            $this->xml = null;
        }

        return $this->fragments;
    }

    /**
     * @param string $text
     *
     * @return $this
     */
    public function addText(string $text): RichText
    {
        if ($this->text) {
            $this->parse($this->text);
        }
        $this->text = $text;
        $this->resetProps();

        return $this;
    }

    /**
     * @return $this
     */
    public function setBold(): RichText
    {
        $this->prop['b'] = true;

        return $this;
    }

    /**
     * @return $this
     */
    public function setItalic(): RichText
    {
        $this->prop['i'] = true;

        return $this;
    }

    /**
     * @return $this
     */
    public function setUnderline(): RichText
    {
        $this->prop['u'] = true;

        return $this;
    }

    /**
     * @param string $font
     *
     * @return $this
     */
    public function setFont(string $font): RichText
    {
        $this->prop['f'] = $font;

        return $this;
    }

    /**
     * @param int $size
     *
     * @return $this
     */
    public function setSize(int $size): RichText
    {
        $this->prop['s'] = $size;

        return $this;
    }

    /**
     * @param string $color
     *
     * @return $this
     */
    public function setColor(string $color): RichText
    {
        $this->prop['c'] = StyleManager::normalizeColor($color);

        return $this;
    }

    /**
     * @return string
     */
    public function outXml(): string
    {
        if ($this->text) {
            $this->parse($this->text);
            $this->text = '';
        }
        if (!$this->xml) {
            $this->xml = '';
            foreach ($this->fragments as $fragment) {
                $rPr = '';
                if ($fragment['prop']['b']) {
                    $rPr .= '<b/>';
                }
                if ($fragment['prop']['i']) {
                    $rPr .= '<i/>';
                }
                if ($fragment['prop']['u']) {
                    $rPr .= '<u/>';
                }
                if ($fragment['prop']['f']) {
                    $rPr .= '<rFont val="' . $fragment['prop']['f'] . '"/>';
                }
                if ($fragment['prop']['s']) {
                    $rPr .= '<sz val="' . $fragment['prop']['s'] . '"/>';
                }
                if ($fragment['prop']['c']) {
                    $rPr .= '<color rgb="' . $fragment['prop']['c'] . '"/>';
                }
                if ($rPr) {
                    $rPr = '<rPr>' . $rPr . '</rPr>';
                }
                $this->xml .= '<r>' . $rPr . '<t xml:space="preserve">' . $fragment['text'] . '</t></r>';
            }
        }

        return $this->xml;
    }
}