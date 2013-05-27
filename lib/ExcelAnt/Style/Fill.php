<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface;

class Fill implements StyleInterface
{
    private $color;

    public function __construct($color = null)
    {
        $this->color = $color ?: 'ff0000';
    }

    public function setColor($color)
    {
        $this->color = $color;

        return $this;
    }

    public function getColor()
    {
        return $this->color;
    }
}