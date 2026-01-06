<?php

namespace avadim\FastExcelWriter\RichText;

use avadim\FastExcelWriter\Style\StyleManager;

class RichTextFragment
{
    protected string $text = '';
    protected int $pos;
    protected array $prop = ['b' => null, 'i' => null, 'u' => null, 'f' => null, 'sz' => null, 'c' => null, 'strike' => null];

    /**
     * RichTextFragment constructor
     *
     * @param string|null $text
     * @param array|null $prop
     */
    public function __construct(?string $text = null, ?array $prop = null)
    {
        $this->text = $text;
        if ($prop) {
            foreach ((array)$prop as $k => $v) {
                $this->prop[$k] = $v;
            }
        }
    }

    protected function setProp(string $key, $value): RichTextFragment
    {
        $this->prop[$key] = $value;

        return $this;
    }

    /**
     * Set font weight to bold
     *
     * @return $this
     */
    public function setBold(): RichTextFragment
    {
        return $this->setProp('b', true);
    }

    /**
     * Set font style to italic
     *
     * @return $this
     */
    public function setItalic(): RichTextFragment
    {
        return $this->setProp('i', true);
    }

    /**
     * Set font decoration to underline
     *
     * @param bool|null $double
     *
     * @return $this
     */
    public function setUnderline(?bool $double = false): RichTextFragment
    {
        return $this->setProp('u', $double ? 'single' : 'double');
    }

    /**
     * Set font decoration to strikethrough
     *
     * @return $this
     */
    public function setStrike(): RichTextFragment
    {
        return $this->setProp('strike', true);
    }

    /**
     * Set font name
     *
     * @param string $font
     *
     * @return $this
     */
    public function setFont(string $font): RichTextFragment
    {
        return $this->setProp('f', $font);
    }

    /**
     * Set font size
     *
     * @param int $size
     *
     * @return $this
     */
    public function setSize(int $size): RichTextFragment
    {
        return $this->setProp('sz', $size);
    }

    /**
     * Set font color
     *
     * @param string $color
     *
     * @return $this
     */
    public function setColor(string $color): RichTextFragment
    {
        return $this->setProp('c', StyleManager::normalizeColor($color));
    }

    /**
     * Get fragment text
     *
     * @return string
     */
    public function getText(): string
    {
        return $this->text;
    }

    /**
     * Converts the object properties into a string representation formatted as XML-like tags.
     *
     * @return string The string representation of the object's properties.
     */
    public function outXml(): string
    {
        $rPr = '';
        if ($this->prop['b']) {
            $rPr .= '<b/>';
        }
        if ($this->prop['i']) {
            $rPr .= '<i/>';
        }
        if ($this->prop['u']) {
            //$rPr .= '<u/>';
            $rPr .= '<u val="' . $this->prop['u'] . '"/>';
        }
        if ($this->prop['strike']) {
            $rPr .= '<strike/>';
        }
        if ($this->prop['f']) {
            $rPr .= '<rFont val="' . $this->prop['f'] . '"/>';
        }
        if ($this->prop['sz']) {
            $rPr .= '<sz val="' . $this->prop['s'] . '"/>';
        }
        if ($this->prop['c']) {
            $rPr .= '<color rgb="' . $this->prop['c'] . '"/>';
        }
        if ($rPr) {
            $rPr = '<rPr>' . $rPr . '</rPr>';
        }

        return '<r>' . $rPr . '<t xml:space="preserve">' . $this->getText() . '</t></r>';
    }
}