<?php

namespace avadim\FastExcelWriter\RichText;

use avadim\FastExcelWriter\Style\StyleManager;
use avadim\FastExcelWriter\Writer\Writer;

class RichTextFragment
{
    protected string $text = '';
    protected int $pos;
    protected array $prop = ['b' => null, 'i' => null, 'u' => null, 'f' => null, 'sz' => null, 'c' => null, 'strike' => null, 'vertAlign' => null];

    /**
     * RichTextFragment constructor
     *
     * @param string|null $text
     * @param array|null $prop
     */
    public function __construct(?string $text = null, ?array $prop = null)
    {
        $this->text = $text ?? '';
        if ($prop) {
            foreach ((array)$prop as $k => $v) {
                $this->prop[$k] = $v;
            }
            if ($this->prop['sz'] !== null) {
                if (!is_numeric($this->prop['sz'])) {
                    throw new \InvalidArgumentException('Font size must be a positive finite number');
                }
                $this->setSize((float)$this->prop['sz']);
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
    public function setBold(bool $enabled = true): RichTextFragment
    {
        return $this->setProp('b', $enabled);
    }

    /** Set subscript for this fragment. */
    public function setSubscript(): RichTextFragment
    {
        return $this->setProp('vertAlign', 'subscript');
    }

    /** Set superscript for this fragment. */
    public function setSuperscript(): RichTextFragment
    {
        return $this->setProp('vertAlign', 'superscript');
    }

    /** Restore the baseline for this fragment. */
    public function setBaseline(): RichTextFragment
    {
        return $this->setProp('vertAlign', 'baseline');
    }

    /**
     * Set font style to italic
     *
     * @return $this
     */
    public function setItalic(bool $enabled = true): RichTextFragment
    {
        return $this->setProp('i', $enabled);
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
        return $this->setProp('u', $double ? 'double' : 'single');
    }

    /** Explicitly disable underline, including inherited formatting. */
    public function removeUnderline(): RichTextFragment
    {
        return $this->setProp('u', 'none');
    }

    /**
     * Set font decoration to strikethrough
     *
     * @return $this
     */
    public function setStrike(bool $enabled = true): RichTextFragment
    {
        return $this->setProp('strike', $enabled);
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
     * @param float $size Positive finite font size in points
     *
     * @return $this
     */
    public function setSize(float $size): RichTextFragment
    {
        if (!is_finite($size) || $size <= 0) {
            throw new \InvalidArgumentException('Font size must be a positive finite number');
        }
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
        if ($this->prop['b'] !== null) {
            $rPr .= $this->prop['b'] ? '<b/>' : '<b val="0"/>';
        }
        if ($this->prop['i'] !== null) {
            $rPr .= $this->prop['i'] ? '<i/>' : '<i val="0"/>';
        }
        if ($this->prop['u']) {
            //$rPr .= '<u/>';
            $rPr .= '<u val="' . Writer::xmlSpecialChars($this->prop['u']) . '"/>';
        }
        if ($this->prop['strike'] !== null) {
            $rPr .= $this->prop['strike'] ? '<strike/>' : '<strike val="0"/>';
        }
        if ($this->prop['f']) {
            $rPr .= '<rFont val="' . Writer::xmlSpecialChars($this->prop['f']) . '"/>';
        }
        if ($this->prop['sz']) {
            $rPr .= '<sz val="' . Writer::floatStr($this->prop['sz']) . '"/>';
        }
        if ($this->prop['c']) {
            $rPr .= '<color rgb="' . Writer::xmlSpecialChars($this->prop['c']) . '"/>';
        }
        if (in_array($this->prop['vertAlign'], ['baseline', 'subscript', 'superscript'], true)) {
            $rPr .= '<vertAlign val="' . $this->prop['vertAlign'] . '"/>';
        }
        if ($rPr) {
            $rPr = '<rPr>' . $rPr . '</rPr>';
        }

        return '<r>' . $rPr . '<t xml:space="preserve">' . Writer::xmlEscapedString($this->getText()) . '</t></r>';
    }
}
