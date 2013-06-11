<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Coordinate\Coordinate;

interface CellInterface
{
    /**
     * Set Coordinate
     *
     * @param Coordinate $coordinate
     *
     * @return CellInterface
     */
    public function setCoordinate(Coordinate $coordinate);

    /**
     * Get Coordinate
     *
     * @return Coordinate
     */
    public function getCoordinate();

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
    public function setStyles(StyleCollection $styles);

    /**
     * Get styles
     *
     * @return StyleCollection
     */
    public function getStyles();

    /**
     * Has styles
     *
     * @return boolean true if there are styles, else false
     */
    public function hasStyles();
}