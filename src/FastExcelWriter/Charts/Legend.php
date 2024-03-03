<?php

namespace avadim\FastExcelWriter\Charts;

class Legend
{
    /** Legend positions */
    const xlLegendPositionBottom = -4107;    //    Below the chart.
    const xlLegendPositionCorner = 2;        //    In the upper right-hand corner of the chart border.
    const xlLegendPositionCustom = -4161;    //    A custom position.
    const xlLegendPositionLeft   = -4131;    //    Left of the chart.
    const xlLegendPositionRight  = -4152;    //    Right of the chart.
    const xlLegendPositionTop    = -4160;    //    Above the chart.

    const POSITION_RIGHT    = 'r';
    const POSITION_LEFT     = 'l';
    const POSITION_BOTTOM   = 'b';
    const POSITION_TOP      = 't';
    const POSITION_TOPRIGHT = 'tr';

    private static array $positionXLref = [
        self::xlLegendPositionBottom => self::POSITION_BOTTOM,
        self::xlLegendPositionCorner => self::POSITION_TOPRIGHT,
        self::xlLegendPositionCustom => '??',
        self::xlLegendPositionLeft   => self::POSITION_LEFT,
        self::xlLegendPositionRight  => self::POSITION_RIGHT,
        self::xlLegendPositionTop    => self::POSITION_TOP
    ];

    /**
     * Legend position
     *
     * @var    string
     */
    private string $position = self::POSITION_RIGHT;

    /**
     * Allow overlay of other elements?
     *
     * @var    boolean
     */
    private bool $overlay = true;

    /**
     * Legend Layout
     *
     * @var Layout|null
     */
    private ?Layout $layout = null;


    /**
     *    Create a new Legend
     */
    public function __construct($position = self::POSITION_RIGHT, Layout $layout = null, $overlay = false)
    {
        $this->setPosition($position);
        $this->layout = $layout;
        $this->setOverlay($overlay);
    }

    /**
     * Get legend position as an excel string value
     *
     * @return    string
     */
    public function getPosition()
    {
        return $this->position;
    }

    /**
     * Get legend position using an excel string value
     *
     * @param    string    $position
     */
    public function setPosition($position = self::POSITION_RIGHT)
    {
        if (!in_array($position, self::$positionXLref)) {
            return false;
        }

        $this->position = $position;
        return true;
    }

    /**
     * Get legend position as an Excel internal numeric value
     *
     * @return    number
     */
    public function getPositionXL()
    {
        return array_search($this->position, self::$positionXLref);
    }

    /**
     * Set legend position using an Excel internal numeric value
     *
     * @param    number    $positionXL
     */
    public function setPositionXL($positionXL = self::xlLegendPositionRight)
    {
        if (!array_key_exists($positionXL, self::$positionXLref)) {
            return false;
        }

        $this->position = self::$positionXLref[$positionXL];
        return true;
    }

    /**
     * Get allow overlay of other elements?
     *
     * @return    boolean
     */
    public function getOverlay()
    {
        return $this->overlay;
    }

    /**
     * Set allow overlay of other elements?
     *
     * @param    boolean    $overlay
     * @return    boolean
     */
    public function setOverlay($overlay = false)
    {
        if (!is_bool($overlay)) {
            return false;
        }

        $this->overlay = $overlay;
        return true;
    }

    /**
     * Get Layout
     *
     * @return Layout
     */
    public function getLayout()
    {
        return $this->layout;
    }

}