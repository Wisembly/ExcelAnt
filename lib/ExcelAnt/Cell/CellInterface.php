<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Collections\StyleCollection;

interface CellInterface
{
    /**
     * Set value
     *
     * @param string $value
     *
     * @return CellInterface
     */
    public function setValue($value);

    /**
     * Get value
     *
     * @return string
     */
    public function getValue();

    /**
     * Set styles
     *
     * @param StyleCollection $styles
     *
     * @return CellInterface
     */
    public function setStyles(StyleCollection $styles = null);

    /**
     * Get styles
     *
     * @return StyleCollection
     */
    public function getStyles();
}