<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Collections\StyleCollection;

class EmptyCell implements CellInterface
{
    /**
     * @throws BadMethodCallException
     */
    public function setValue($value)
    {
        throw new \BadMethodCallException("This method cannot be called");
    }

    /**
     * @throws BadMethodCallException
     */
    public function getValue()
    {
        throw new \BadMethodCallException("This method cannot be called");
    }

    /**
     * @throws BadMethodCallException
     */
    public function setStyles(StyleCollection $styles)
    {
        throw new \BadMethodCallException("This method cannot be called");
    }

    /**
     * @throws BadMethodCallException
     */
    public function getStyles()
    {
        throw new \BadMethodCallException("This method cannot be called");
    }
}