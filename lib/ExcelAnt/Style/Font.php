<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface;

class Font implements StyleInterface
{
    /* Underline types */
    const UNDERLINE_NONE             = 'none';
    const UNDERLINE_DOUBLE           = 'double';
    const UNDERLINE_DOUBLEACCOUNTING = 'doubleAccounting';
    const UNDERLINE_SINGLE           = 'single';
    const UNDERLINE_SINGLEACCOUNTING = 'singleAccounting';

    private $name = 'Calibri';
    private $size = 11;
    private $bold = false;
    private $italic = false;
    private $color = '000000';
    private $underline = self::UNDERLINE_NONE;

    /**
     * Set name
     *
     * @param string $name
     *
     * @return Font
     */
    public function setName($name)
    {
        $this->name = $name;

        return $this;
    }

    /**
     * Get name
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Set size
     *
     * @param string $size
     *
     * @throws InvalidArgumentException If size isn't numeric value
     *
     * @return Font
     */
    public function setSize($size)
    {
        if (false === filter_var($size, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("The parameter must be a numeric value");
        }

        $this->size = $size;

        return $this;
    }

    /**
     * Get size
     *
     * @return string
     */
    public function getSize()
    {
        return $this->size;
    }

    /**
     * Set bold
     *
     * @param string $bold
     *
     * @throws InvalidArgumentException If bold isn't boolean
     *
     * @return Font
     */
    public function setBold($bold)
    {
        if (!is_bool($bold)) {
            throw new \InvalidArgumentException("The parameter must be a boolean value");
        }

        $this->bold = $bold;

        return $this;
    }

    /**
     * If the font is bold
     *
     * @return boolean
     */
    public function isBold()
    {
        return $this->bold;
    }

    /**
     * Set italic
     *
     * @param string $italic
     *
     * @throws InvalidArgumentException If italic isn't boolean
     *
     * @return Font
     */
    public function setItalic($italic)
    {
        if (!is_bool($italic)) {
            throw new \InvalidArgumentException("The parameter must be a boolean value");
        }

        $this->italic = $italic;

        return $this;
    }

    /**
     * If is italic
     *
     * @return boolean
     */
    public function isItalic()
    {
        return $this->italic;
    }

    /**
     * Set color
     *
     * @param string $color
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
     * Set underline
     *
     * @param string $underline
     *
     * @throws InvalidArgumentException If underline doesn't belong the underline parameter list
     *
     * @return Font
     */
    public function setUnderline($underline)
    {
        if (!in_array($underline, $this->getUnderlines())) {
            throw new \InvalidArgumentException("The parameter must belong to the Underline parameter list");
        }

        $this->underline = $underline;

        return $this;
    }

    /**
     * Get underline
     *
     * @return string
     */
    public function getUnderline()
    {
        return $this->underline;
    }

    /**
     * Return the underline parameter list
     *
     * @return array
     */
    public function getUnderlines()
    {
        return [
            self::UNDERLINE_NONE,
            self::UNDERLINE_DOUBLE,
            self::UNDERLINE_DOUBLEACCOUNTING,
            self::UNDERLINE_SINGLE,
            self::UNDERLINE_SINGLEACCOUNTING,
        ];
    }
}