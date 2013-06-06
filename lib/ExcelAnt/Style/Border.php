<?php

namespace ExcelAnt\Style;

class Border
{
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

    private $color = '000000';
    private $type = self::BORDER_THIN;

    /**
     * Set type
     *
     * @throws InvalidArgumentException If the parameter doesn't belong the type parameter list
     *
     * @param string $type
     *
     * @return Border
     */
    public function setType($type)
    {
        if (!in_array($type, $this->getTypes())) {
            throw new \InvalidArgumentException("The parameter must belong to the Type parameter list");
        }

        $this->type = $type;

        return $this;
    }

    /**
     * Get type
     *
     * @return Border
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * Set color
     *
     * @param string $color
     *
     * @return Border
     */
    public function setColor($color)
    {
        $this->color = $color;

        return $this;
    }

    /**
     * Get color
     *
     * @return string
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * Return the type parameter list
     *
     * @return array
     */
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
}