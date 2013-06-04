<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface;

class Alignment implements StyleInterface
{
    /* Horizontal alignment styles */
    const HORIZONTAL_GENERAL                = 'general';
    const HORIZONTAL_LEFT                   = 'left';
    const HORIZONTAL_RIGHT                  = 'right';
    const HORIZONTAL_CENTER                 = 'center';
    const HORIZONTAL_CENTER_CONTINUOUS      = 'centerContinuous';
    const HORIZONTAL_JUSTIFY                = 'justify';

    /* Vertical alignment styles */
    const VERTICAL_BOTTOM                   = 'bottom';
    const VERTICAL_TOP                      = 'top';
    const VERTICAL_CENTER                   = 'center';
    const VERTICAL_JUSTIFY                  = 'justify';

    private $horizontal = self::HORIZONTAL_GENERAL;
    private $vertical = self::VERTICAL_BOTTOM;

    /**
     * Set vertical alignment
     *
     * @param string $alignment
     *
     * @throws InvalidArgumentException If the parameter doesn't belong the vertical parameter list
     *
     * @return Alignment
     */
    public function setVertical($alignment)
    {
        $this->checkVerticalParameter($alignment);
        $this->vertical = $alignment;

        return $this;
    }

    /**
     * Get vertical alignment
     *
     * @return string
     */
    public function getVertical()
    {
        return $this->vertical;
    }

    /**
     * Set horizontal alignment
     *
     * @param string $alignment
     *
     * @throws InvalidArgumentException If the parameter doesn't belong the horizontal parameter list
     *
     * @return Alignment
     */
    public function setHorizontal($alignment)
    {
        $this->checkHorizontalParameter($alignment);
        $this->horizontal = $alignment;

        return $this;
    }

    /**
     * Get horizontal alignment
     *
     * @return string
     */
    public function getHorizontal()
    {
        return $this->horizontal;
    }

    /**
     * Return the horizontal parameter list
     *
     * @return array
     */
    public function getHorizontals()
    {
        return [
            self::HORIZONTAL_GENERAL,
            self::HORIZONTAL_LEFT,
            self::HORIZONTAL_RIGHT,
            self::HORIZONTAL_CENTER,
            self::HORIZONTAL_CENTER_CONTINUOUS,
            self::HORIZONTAL_JUSTIFY,
        ];
    }

    /**
     * Return the vertical parameter list
     *
     * @return array
     */
    public function getVerticals()
    {
        return [
            self::VERTICAL_BOTTOM,
            self::VERTICAL_TOP,
            self::VERTICAL_CENTER,
            self::VERTICAL_JUSTIFY,
        ];
    }

    /**
     * Check if the alignment parameter belong the horizontal parameter list
     *
     * @param  string $alignment
     *
     * @throws InvalidArgumentException If the parameter doesn't belong the horizontal parameter list
     */
    private function checkHorizontalParameter($alignment)
    {
        if (!in_array($alignment, $this->getHorizontals())) {
            throw new \InvalidArgumentException("The parameter must belong to the horizontal parameter list");
        }
    }

    /**
     * Check if the alignment parameter belong the vertical parameter list
     *
     * @param  string $alignment
     *
     * @throws InvalidArgumentException If the parameter doesn't belong the vertical parameter list
     */
    private function checkVerticalParameter($alignment)
    {
        if (!in_array($alignment, $this->getVerticals())) {
            throw new \InvalidArgumentException("The parameter must belong to the vertical parameter list");
        }
    }
}