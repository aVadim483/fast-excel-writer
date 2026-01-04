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
    public function setLayoutTarget(string $value): self
    {
        $this->layoutTarget = $value;
        
        return $this;
    }

    /**
     * Get X-Mode
     *
     * @return string
     */
    public function getXMode(): string
    {
        return $this->xMode;
    }

    /**
     * Set X-Mode
     *
     * @param string $value
     * 
     * @return $this
     */
    public function setXMode(string $value): self
    {
        $this->xMode = $value;

        return $this;
    }

    /**
     * Get Y-Mode
     *
     * @return string
     */
    public function getYMode(): string
    {
        return $this->yMode;
    }

    /**
     * Set Y-Mode
     *
     * @param string $value
     *
     * @return $this
     */
    public function setYMode(string $value): Layout
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
    public function setXPosition(float $value): Layout
    {
        $this->xPos = $value;
        
        return $this;
    }

    /**
     * Get Y-Position
     *
     * @return float|int
     */
    public function getYPosition()
    {
        return $this->yPos;
    }

    /**
     * Set Y-Position
     *
     * @param $value
     * 
     * @return $this
     */
    public function setYPosition($value): Layout
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
     * 
     * @return $this
     */
    public function setWidth($value): Layout
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
     * 
     * @return $this
     */
    public function setHeight(float $value): Layout
    {
        $this->height = $value;
        
        return $this;
    }


    /**
     * Get show legend key
     *
     * @return bool
     */
    public function getShowLegendKey(): bool
    {
        return $this->showLegendKey;
    }

    /**
     * Set show legend key
     * Specifies that legend keys should be shown in data labels.
     *
     * @param bool $value Show legend key
     * 
     * @return $this
     */
    public function setShowLegendKey(bool $value): Layout
    {
        $this->showLegendKey = $value;
        return $this;
    }

    /**
     * Get show value
     *
     * @return bool
     */
    public function getShowVal(): bool
    {
        return $this->showVal;
    }

    /**
     * Set show val
     * Specifies that the value should be shown in data labels.
     *
     * @param bool $value Show val
     *
     * @return $this
     */
    public function setShowVal(bool $value): Layout
    {
        $this->showVal = $value;
        return $this;
    }

    /**
     * Get show category name
     *
     * @return bool
     */
    public function getShowCatName(): bool
    {
        return $this->showCatName;
    }

    /**
     * Set show cat name
     * Specifies that the category name should be shown in data labels.
     *
     * @param bool $value        Show cat name
     *
     * @return $this
     */
    public function setShowCatName(bool $value): Layout
    {
        $this->showCatName = $value;

        return $this;
    }

    /**
     * Get show data series name
     *
     * @return bool
     */
    public function getShowSerName(): bool
    {
        return $this->showSerName;
    }

    /**
     * Set show ser name
     * Specifies that the series name should be shown in data labels.
     *
     * @param bool $value Show series name
     *
     * @return $this
     */
    public function setShowSerName(bool $value): Layout
    {
        $this->showSerName = $value;

        return $this;
    }

    /**
     * Get show percentage
     *
     * @return bool
     */
    public function getShowPercent(): bool
    {
        return $this->showPercent;
    }

    /**
     * Set show percentage
     * Specifies that the percentage should be shown in data labels.
     *
     * @param bool $value Show percentage
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
     * @return bool
     */
    public function getShowBubbleSize(): bool
    {
        return $this->showBubbleSize;
    }

    /**
     * Set show bubble size
     * Specifies that the bubble size should be shown in data labels.
     *
     * @param bool $value Show bubble size
     *
     * @return $this
     */
    public function setShowBubbleSize(bool $value): Layout
    {
        $this->showBubbleSize = $value;

        return $this;
    }

    /**
     * Get show leader lines
     *
     * @return bool
     */
    public function getShowLeaderLines(): bool
    {
        return $this->showLeaderLines;
    }

    /**
     * Set show leader lines
     * Specifies that leader lines should be shown in data labels.
     *
     * @param bool $value Show leader lines
     *
     * @return $this
     */
    public function setShowLeaderLines(bool $value): Layout
    {
        $this->showLeaderLines = $value;

        return $this;
    }

}