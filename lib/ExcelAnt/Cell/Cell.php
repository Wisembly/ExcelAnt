<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Collections\StyleCollection;
use ExcelAnt\Traits\Coordinable;

class Cell implements CellInterface
{
    use Coordinable;

    private $coordinate;
    private $value;
    private $styleCollection;

    /**
     * @param string $value
     */
    public function __construct($value = null)
    {
        if (isset($value)) {
            $this->value = $value;
        }
    }

    /**
     * {@inheritdoc}
     */
    public function setValue($value)
    {
        $this->value = $value;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getValue()
    {
        return $this->value;
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
}