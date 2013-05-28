<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface;

class Fill implements StyleInterface
{
    private $color;

    public function __construct($color = null)
    {
        $this->color = $color ?: 'ffffff';
    }

    /**
     * Set color
     *
     * @param string $color
     *
     * @return Fill
     */
    public function setColor($color)
    {
        $this->color = $color;

        return $this;
    }

    /**
     * Get color
     *
     * @return string
     */
    public function getColor()
    {
        return $this->color;
    }
}