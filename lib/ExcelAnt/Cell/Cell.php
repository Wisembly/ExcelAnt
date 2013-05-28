<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Style\StyleInterface;

class Cell implements CellInterface
{
    public function __construct()
    {

    }

    public function setValue($value)
    {
        $this->value = $value;

        return $this;
    }

    public function getValue()
    {
        return $this->value;
    }

    public function setStyle(StyleInterface $style)
    {
        $this->style = $style;

        return $this;
    }

    public function getStyle()
    {
        return $this->style;
    }
}