<?php

namespace avadim\FastExcelWriter\Charts;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
class Layout
{
    /**
     * layoutTarget
     *
     * @var string
     */
    private $layoutTarget;

    /**
     * X Mode
     *
     * @var string
     */
    private $xMode;

    /**
     * Y Mode
     *
     * @var string
     */
    private $yMode;

    /**
     * X-Position
     *
     * @var float
     */
    private float $xPos = 0;

    /**
     * Y-Position
     *
     * @var float
     */
    private float $yPos = 0;

    /**
     * width
     *
     * @var float
     */
    private float $width = 0;

    /**
     * height
     *
     * @var float
     */
    private float $height = 0;

    /**
     * show legend key
     * Specifies that legend keys should be shown in data labels
     *
     * @var boolean
     */
    private bool $showLegendKey = false;

    /**
     * show value
     * Specifies that the value should be shown in a data label.
     *
     * @var boolean
     */
    private bool $showVal = false;

    /**
     * show category name
     * Specifies that the category name should be shown in the data label.
     *
     * @var boolean
     */
    private bool $showCatName = false;

    /**
     * show data series name
     * Specifies that the series name should be shown in the data label.
     *
     * @var boolean
     */
    private bool $showSerName = false;

    /**
     * show percentage
     * Specifies that the percentage should be shown in the data label.
     *
     * @var boolean
     */
    private bool $showPercent = false;

    /**
     * show bubble size
     *
     * @var boolean
     */
    private bool $showBubbleSize = false;

    /**
     * show leader lines
     * Specifies that leader lines should be shown for the data label.
     *
     * @var boolean
     */
    private bool $showLeaderLines = false;


    /**
     * Create a new $this
     */
    public function __construct($layout = [])
    {
        if (isset($layout['layoutTarget'])) {
            $this->layoutTarget = $layout['layoutTarget'];
        }
        if (isset($layout['xMode'])) {
            $this->xMode = $layout['xMode'];
        }
        if (isset($layout['yMode'])) {
            $this->yMode = $layout['yMode'];
        }
        if (isset($layout['x'])) {
            $this->xPos = (float) $layout['x'];
        }
        if (isset($layout['y'])) {
            $this->yPos = (float) $layout['y'];
        }
        if (isset($layout['w'])) {
            $this->width = (float) $layout['w'];
        }
        if (isset($layout['h'])) {
            $this->height = (float) $layout['h'];
        }
    }

    /**
     * Get Layout Target
     *
     * @return string|null
     */
    public function getLayoutTarget(): ?string
    {
        return $this->layoutTarget;
    }

    /**
     * Set Layout Target
     *
     * @param string $value Target value
     *
     * @return $this
     */
    public function setLayoutTarget(string $value)
    {
        $this->layoutTarget = $value;
        return $this;
    }

    /**
     * Get X-Mode
     *
     * @return string
     */
    public function getXMode()
    {
        return $this->xMode;
    }

    /**
     * Set X-Mode
     *
     * @param $value
     * @return $this
     */
    public function setXMode($value)
    {
        $this->xMode = $value;
        return $this;
    }

    /**
     * Get Y-Mode
     *
     * @return string
     */
    public function getYMode()
    {
        return $this->yMode;
    }

    /**
     * Set Y-Mode
     *
     * @param $value
     * @return $this
     */
    public function setYMode($value)
    {
        $this->yMode = $value;
        return $this;
    }

    /**
     * Get X-Position
     *
     * @return float
     */
    public function getXPosition()
    {
        return $this->xPos;
    }

    /**
     * Set X-Position
     *
     * @param float $value
     *
     * @return $this
     */
    public function setXPosition(float $value)
    {
        $this->xPos = $value;
        return $this;
    }

    /**
     * Get Y-Position
     *
     * @return number
     */
    public function getYPosition()
    {
        return $this->yPos;
    }

    /**
     * Set Y-Position
     *
     * @param $value
     * @return $this
     */
    public function setYPosition($value)
    {
        $this->yPos = $value;
        return $this;
    }

    /**
     * Get Width
     *
     * @return float
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set Width
     *
     * @param $value
     * @return $this
     */
    public function setWidth($value)
    {
        $this->width = $value;
        return $this;
    }

    /**
     * Get Height
     *
     * @return float
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Set Height
     *
     * @param float $value
     * @return $this
     */
    public function setHeight($value)
    {
        $this->height = $value;
        return $this;
    }


    /**
     * Get show legend key
     *
     * @return boolean
     */
    public function getShowLegendKey()
    {
        return $this->showLegendKey;
    }

    /**
     * Set show legend key
     * Specifies that legend keys should be shown in data labels.
     *
     * @param boolean $value        Show legend key
     * @return $this
     */
    public function setShowLegendKey($value)
    {
        $this->showLegendKey = $value;
        return $this;
    }

    /**
     * Get show value
     *
     * @return boolean
     */
    public function getShowVal()
    {
        return $this->showVal;
    }

    /**
     * Set show val
     * Specifies that the value should be shown in data labels.
     *
     * @param boolean $value        Show val
     * @return $this
     */
    public function setShowVal($value)
    {
        $this->showVal = $value;
        return $this;
    }

    /**
     * Get show category name
     *
     * @return boolean
     */
    public function getShowCatName()
    {
        return $this->showCatName;
    }

    /**
     * Set show cat name
     * Specifies that the category name should be shown in data labels.
     *
     * @param boolean $value        Show cat name
     * @return $this
     */
    public function setShowCatName($value)
    {
        $this->showCatName = $value;
        return $this;
    }

    /**
     * Get show data series name
     *
     * @return boolean
     */
    public function getShowSerName()
    {
        return $this->showSerName;
    }

    /**
     * Set show ser name
     * Specifies that the series name should be shown in data labels.
     *
     * @param boolean $value        Show series name
     * @return $this
     */
    public function setShowSerName($value)
    {
        $this->showSerName = $value;
        return $this;
    }

    /**
     * Get show percentage
     *
     * @return boolean
     */
    public function getShowPercent(): bool
    {
        return $this->showPercent;
    }

    /**
     * Set show percentage
     * Specifies that the percentage should be shown in data labels.
     *
     * @param boolean $value Show percentage
     *
     * @return $this
     */
    public function setShowPercent(bool $value): Layout
    {
        $this->showPercent = $value;
        return $this;
    }

    /**
     * Get show bubble size
     *
     * @return boolean
     */
    public function getShowBubbleSize()
    {
        return $this->showBubbleSize;
    }

    /**
     * Set show bubble size
     * Specifies that the bubble size should be shown in data labels.
     *
     * @param boolean $value        Show bubble size
     * @return $this
     */
    public function setShowBubbleSize($value)
    {
        $this->showBubbleSize = $value;
        return $this;
    }

    /**
     * Get show leader lines
     *
     * @return boolean
     */
    public function getShowLeaderLines()
    {
        return $this->showLeaderLines;
    }

    /**
     * Set show leader lines
     * Specifies that leader lines should be shown in data labels.
     *
     * @param boolean $value        Show leader lines
     * @return $this
     */
    public function setShowLeaderLines($value)
    {
        $this->showLeaderLines = $value;
        return $this;
    }

}