<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\CellInterface,
    ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Traits\Coordinable;

class EmptyCell implements CellInterface
{
    use Coordinable;

    private $styleCollection;

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
     * {@inheritdoc}
     */
    public function setStyles(StyleCollection $styles)
    {
        $this->styleCollection = $styles;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getStyles()
    {
        return $this->styleCollection;
    }

    /**
     * {@inheritdoc}
     */
    public function hasStyles()
    {
        return empty($this->styleCollection) ? false : true;
    }
}