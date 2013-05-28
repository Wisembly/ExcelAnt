<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Style\StyleInterface;

class Cell implements CellInterface
{
    use StyleCollectionRequirements;

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
    public function setStyle($styles = null)
    {
        $this->styleCollection = $style;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getStyle()
    {
        return $this->style;
    }
}