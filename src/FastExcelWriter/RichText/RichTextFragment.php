<?php

namespace avadim\FastExcelWriter\RichText;

use avadim\FastExcelWriter\StyleManager;

class RichTextFragment
{
    protected string $text = '';
    protected int $pos;
    protected array $prop = ['b' => null, 'i' => null, 'u' => null, 'f' => null, 's' => null, 'c' => null];

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

    /**
     * Set font weight to bold
     *
     * @return $this
     */
    public function setBold(): RichTextFragment
    {
        $this->prop['b'] = true;

        return $this;
    }

    /**
     * Set font style to italic
     *
     * @return $this
     */
    public function setItalic(): RichTextFragment
    {
        $this->prop['i'] = true;

        return $this;
    }

    /**
     * Set font decoration to underline
     *
     * @return $this
     */
    public function setUnderline(): RichTextFragment
    {
        $this->prop['u'] = true;

        return $this;
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
        $this->prop['f'] = $font;

        return $this;
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
        $this->prop['s'] = $size;

        return $this;
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
        $this->prop['c'] = StyleManager::normalizeColor($color);

        return $this;
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
            $rPr .= '<u/>';
        }
        if ($this->prop['f']) {
            $rPr .= '<rFont val="' . $this->prop['f'] . '"/>';
        }
        if ($this->prop['s']) {
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