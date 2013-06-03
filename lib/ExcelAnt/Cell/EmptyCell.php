<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Collections\StyleCollection;
use ExcelAnt\Traits\Coordinable;

class EmptyCell implements CellInterface
{
    use Coordinable;

    private $coordinate;

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

    /**
     * {@inheritdoc}
     */
    public function hasStyles()
    {
        throw new \BadMethodCallException("This method cannot be called");
    }
}