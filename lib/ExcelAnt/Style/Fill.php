<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface;

class Fill implements StyleInterface
{
    /* Fill types */
    const FILL_NONE                         = 'none';
    const FILL_SOLID                        = 'solid';
    const FILL_GRADIENT_LINEAR              = 'linear';
    const FILL_GRADIENT_PATH                = 'path';
    const FILL_PATTERN_DARKDOWN             = 'darkDown';
    const FILL_PATTERN_DARKGRAY             = 'darkGray';
    const FILL_PATTERN_DARKGRID             = 'darkGrid';
    const FILL_PATTERN_DARKHORIZONTAL       = 'darkHorizontal';
    const FILL_PATTERN_DARKTRELLIS          = 'darkTrellis';
    const FILL_PATTERN_DARKUP               = 'darkUp';
    const FILL_PATTERN_DARKVERTICAL         = 'darkVertical';
    const FILL_PATTERN_GRAY0625             = 'gray0625';
    const FILL_PATTERN_GRAY125              = 'gray125';
    const FILL_PATTERN_LIGHTDOWN            = 'lightDown';
    const FILL_PATTERN_LIGHTGRAY            = 'lightGray';
    const FILL_PATTERN_LIGHTGRID            = 'lightGrid';
    const FILL_PATTERN_LIGHTHORIZONTAL      = 'lightHorizontal';
    const FILL_PATTERN_LIGHTTRELLIS         = 'lightTrellis';
    const FILL_PATTERN_LIGHTUP              = 'lightUp';
    const FILL_PATTERN_LIGHTVERTICAL        = 'lightVertical';
    const FILL_PATTERN_MEDIUMGRAY           = 'mediumGray';

    private $color = 'ffffff';
    private $type = self::FILL_SOLID;

    /**
     * Set color
     *
     * @param string $color
     *
     * @return Fill
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
     * Set type
     *
     * @throws InvalidArgumentException If underline doesn't belong the underline parameter list
     *
     * @param string $type
     *
     * @return Fill
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
     * @return string
     */
    public function getType()
    {
        return $this->type;
    }

    public function getTypes()
    {
        return [
            self::FILL_NONE,
            self::FILL_SOLID,
            self::FILL_GRADIENT_LINEAR,
            self::FILL_GRADIENT_PATH,
            self::FILL_PATTERN_DARKDOWN,
            self::FILL_PATTERN_DARKGRAY,
            self::FILL_PATTERN_DARKGRID,
            self::FILL_PATTERN_DARKHORIZONTAL,
            self::FILL_PATTERN_DARKTRELLIS,
            self::FILL_PATTERN_DARKUP,
            self::FILL_PATTERN_DARKVERTICAL,
            self::FILL_PATTERN_GRAY0625,
            self::FILL_PATTERN_GRAY125,
            self::FILL_PATTERN_LIGHTDOWN,
            self::FILL_PATTERN_LIGHTGRAY,
            self::FILL_PATTERN_LIGHTGRID,
            self::FILL_PATTERN_LIGHTHORIZONTAL,
            self::FILL_PATTERN_LIGHTTRELLIS,
            self::FILL_PATTERN_LIGHTUP,
            self::FILL_PATTERN_LIGHTVERTICAL,
            self::FILL_PATTERN_MEDIUMGRAY,
        ];
    }
}