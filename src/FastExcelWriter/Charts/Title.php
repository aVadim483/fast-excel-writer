<?php

namespace avadim\FastExcelWriter\Charts;

/**
 * This class uses source code of PHPExcel
 *
 * @license LGPL http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
 */
class Title
{
    /**
     * Title Caption
     *
     * @var string|null
     */
    private ?string $caption = null;

    /**
     * Title Layout
     *
     * @var Layout|null
     */
    private ?Layout $layout = null;

    /**
     * Create a new Title
     */
    public function __construct(?string $caption = null, ?Layout $layout = null)
    {
        $this->caption = $caption;
        $this->layout = $layout;
    }

    /**
     * Get caption
     *
     * @return string
     */
    public function getCaption(): ?string
    {
        return $this->caption;
    }

    /**
     * Set caption
     *
     * @param string|null $caption
     *
     * @return $this
     */
    public function setCaption(?string $caption = null): Title
    {
        $this->caption = $caption;

        return $this;
    }

    /**
     * Get Layout
     *
     * @return Layout|null
     */
    public function getLayout(): ?Layout
    {
        return $this->layout;
    }

}