<?php

namespace avadim\FastExcelWriter\Charts;

class Title
{
    /**
     * Title Caption
     *
     * @var string
     */
    private $caption = null;

    /**
     * Title Layout
     *
     * @var Layout|null
     */
    private ?Layout $layout = null;

    /**
     * Create a new Title
     */
    public function __construct($caption = null, Layout $layout = null)
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