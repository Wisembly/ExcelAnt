<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface;

class Border implements StyleInterface
{
    /* Border side */
    const SIDE_LEFT   = 'left';
    const SIDE_TOP    = 'top';
    const SIDE_RIGHT  = 'right';
    const SIDE_BOTTOM = 'bottom';
    const SIDE_ALL    = 'all';

    /* Border style */
    const BORDER_NONE             = 'none';
    const BORDER_DASHDOT          = 'dashDot';
    const BORDER_DASHDOTDOT       = 'dashDotDot';
    const BORDER_DASHED           = 'dashed';
    const BORDER_DOTTED           = 'dotted';
    const BORDER_DOUBLE           = 'double';
    const BORDER_HAIR             = 'hair';
    const BORDER_MEDIUM           = 'medium';
    const BORDER_MEDIUMDASHDOT    = 'mediumDashDot';
    const BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot';
    const BORDER_MEDIUMDASHED     = 'mediumDashed';
    const BORDER_SLANTDASHDOT     = 'slantDashDot';
    const BORDER_THICK            = 'thick';
    const BORDER_THIN             = 'thin';

    private $side;
    private $color;
    private $type;

    public function __construct($side)
    {
        $this->checkSideParameter($side);
        $this->side = $side;
        $this->type = self::BORDER_NONE;
        $this->color = '000000';
    }

    public function setSide($side)
    {
        $this->checkSideParameter($side);
        $this->side = $side;

        return $this;
    }

    public function getSide()
    {
        return $this->side;
    }

    public function setType($type)
    {
        if (!in_array($type, $this->getTypes())) {
            throw new \InvalidArgumentException("The parameter must belong to the Type parameter list");
        }

        $this->type = $type;

        return $this;
    }

    public function getType()
    {
        return $this->type;
    }

    public function setColor($color)
    {
        $this->color = $color;

        return $this;
    }

    public function getColor()
    {
        return $this->color;
    }

    public function getSides()
    {
        return [
            self::SIDE_LEFT,
            self::SIDE_TOP,
            self::SIDE_RIGHT,
            self::SIDE_BOTTOM,
        ];
    }

    public function getTypes()
    {
        return [
            self::BORDER_NONE,
            self::BORDER_DASHDOT,
            self::BORDER_DASHDOTDOT,
            self::BORDER_DASHED,
            self::BORDER_DOTTED,
            self::BORDER_DOUBLE,
            self::BORDER_HAIR,
            self::BORDER_MEDIUM,
            self::BORDER_MEDIUMDASHDOT,
            self::BORDER_MEDIUMDASHDOTDOT,
            self::BORDER_MEDIUMDASHED,
            self::BORDER_SLANTDASHDOT,
            self::BORDER_THICK,
            self::BORDER_THIN,
        ];
    }

    private function checkSideParameter($side)
    {
        if (!in_array($side, $this->getSides())) {
            throw new \InvalidArgumentException("The parameter must belong to the Side parameter list");
        }
    }
}