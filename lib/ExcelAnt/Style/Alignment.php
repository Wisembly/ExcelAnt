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

    private $horizontal;
    private $vertical;

    public function setVertical($alignment)
    {
        $this->checkVerticalParameter($alignment);
        $this->vertical = $alignment;
        $this->horizontal = self::HORIZONTAL_GENERAL;
        $this->vertical = self::VERTICAL_BOTTOM;

        return $this;
    }

    public function getVertical()
    {
        return $this->vertical;
    }

    public function setHorizontal($alignment)
    {
        $this->checkHorizontalParameter($alignment);
        $this->horizontal = $alignment;

        return $this;
    }

    public function getHorizontal()
    {
        return $this->horizontal;
    }

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

    public function getVerticals()
    {
        return [
            self::VERTICAL_BOTTOM,
            self::VERTICAL_TOP,
            self::VERTICAL_CENTER,
            self::VERTICAL_JUSTIFY,
        ];
    }

    private function checkHorizontalParameter($alignment)
    {
        if (!in_array($alignment, $this->getHorizontals())) {
            throw new \InvalidArgumentException("The parameter must belong to the Horizontal parameter list");
        }
    }

    private function checkVerticalParameter($alignment)
    {
        if (!in_array($alignment, $this->getVerticals())) {
            throw new \InvalidArgumentException("The parameter must belong to the Vertical parameter list");
        }
    }
}