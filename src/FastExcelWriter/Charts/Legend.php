<?php

namespace avadim\FastExcelWriter\Charts;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
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
    public function __construct($position = self::POSITION_RIGHT, ?Layout $layout = null, $overlay = false)
    {
        $this->setPosition($position);
        $this->layout = $layout;
        $this->setOverlay($overlay);
    }

    /**
     * Get legend position as an excel string value
     *
     * @return string
     */
    public function getPosition(): string
    {
        return $this->position;
    }

    /**
     * Get legend position using an excel string value
     *
     * @param string|null $position
     *
     * @return bool
     */
    public function setPosition(?string $position = self::POSITION_RIGHT): bool
    {
        if (!in_array($position, self::$positionXLref)) {
            return false;
        }
        $this->position = $position;

        return true;
    }

    /**
     * @return bool
     */
    public function setPositionTop(): bool
    {
        return $this->setPosition(self::POSITION_TOP);
    }

    /**
     * @return bool
     */
    public function setPositionRight(): bool
    {
        return $this->setPosition(self::POSITION_RIGHT);
    }

    /**
     * @return bool
     */
    public function setPositionBottom(): bool
    {
        return $this->setPosition(self::POSITION_BOTTOM);
    }

    /**
     * @return bool
     */
    public function setPositionLeft(): bool
    {
        return $this->setPosition(self::POSITION_LEFT);
    }

    /**
     * Get allow overlay of other elements?
     *
     * @return bool
     */
    public function getOverlay(): bool
    {
        return $this->overlay;
    }

    /**
     * Set allow overlay of other elements?
     *
     * @param bool $overlay
     *
     * @return bool
     */
    public function setOverlay(bool $overlay = false): bool
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
    public function getLayout(): ?Layout
    {
        return $this->layout;
    }

}