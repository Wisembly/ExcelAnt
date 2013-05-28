<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Style\StyleInterface;

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
     * Set style
     *
     * @param StyleInterface $style
     *
     * @return CellInterface
     */
    public function setStyle(StyleInterface $style);

    /**
     * Get style
     *
     * @return StyleInterface
     */
    public function getStyle();
}