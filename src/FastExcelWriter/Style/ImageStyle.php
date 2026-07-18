<?php

namespace avadim\FastExcelWriter\Style;

/**
 * Class ImageStyle
 * @package avadim\FastExcelWriter\Style
 */
class ImageStyle
{
    /** @var int|float|null */
    public $width = null;

    /** @var int|float|null */
    public $height = null;

    /** @var int|float|null */
    public $x = null;

    /** @var int|float|null */
    public $y = null;

    /** @var string|null */
    public ?string $hyperlink = null;


    /**
     * ImageStyle constructor
     *
     * @param array $options
     */
    public function __construct(array $options = [])
    {
        if ($options) {
            $this->set($options);
        }
    }

    /**
     * Set style options from array
     *
     * @param array $options
     * @return $this
     */
    public function set(array $options): ImageStyle
    {
        foreach ($options as $key => $value) {
            if (property_exists($this, $key)) {
                $this->$key = $value;
            }
        }
        return $this;
    }

    /**
     * Set width of image
     *
     * @param int|float $width
     * @return $this
     */
    public function width($width): ImageStyle
    {
        $this->width = $width;

        return $this;
    }

    /**
     * Set height of image
     *
     * @param int|float $height
     * @return $this
     */
    public function height($height): ImageStyle
    {
        $this->height = $height;

        return $this;
    }

    /**
     * Set offset in pixels relative to the left and top borders of the cell
     *
     * @param int|float $x
     * @param int|float $y
     * @return $this
     */
    public function offset($x, $y): ImageStyle
    {
        $this->x = $x;
        $this->y = $y;

        return $this;
    }

    /**
     * Set URL of hyperlink
     *
     * @param string $hyperlink
     * @return $this
     */
    public function hyperlink(string $hyperlink): ImageStyle
    {
        $this->hyperlink = $hyperlink;

        return $this;
    }

    /**
     * Return style options as array
     *
     * @return array
     */
    public function toArray(): array
    {
        $result = [];
        if ($this->width !== null) {
            $result['width'] = $this->width;
        }
        if ($this->height !== null) {
            $result['height'] = $this->height;
        }
        if ($this->x !== null) {
            $result['x'] = $this->x;
        }
        if ($this->y !== null) {
            $result['y'] = $this->y;
        }
        if ($this->hyperlink !== null) {
            $result['hyperlink'] = $this->hyperlink;
        }
        return $result;
    }
}
