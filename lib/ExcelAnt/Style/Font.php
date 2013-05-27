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

    private $name;
    private $size;
    private $bold;
    private $italic;
    private $color;
    private $underline;

    public function __construct()
    {
        $this->name = 'Calibri';
        $this->size = 11;
        $this->bold = false;
        $this->italic = false;
        $this->color = '000000';
        $this->underline = self::UNDERLINE_NONE;
    }

    public function setName($name)
    {
        $this->name = $name;

        return $this;
    }

    public function getName()
    {
        return $this->name;
    }

    public function setSize($size)
    {
        if (!is_numeric($size)) {
            throw new \InvalidArgumentException("The parameter must be a numeric value");
        }

        $this->size = $size;

        return $this;
    }

    public function getSize()
    {
        return $this->size;
    }

    public function setBold($bold)
    {
        if (!is_bool($bold)) {
            throw new \InvalidArgumentException("The parameter must be a boolean value");
        }

        $this->bold = $bold;

        return $this;
    }

    public function isBold()
    {
        return $this->bold;
    }

    public function setItalic($italic)
    {
        if (!is_bool($italic)) {
            throw new \InvalidArgumentException("The parameter must be a boolean value");
        }

        $this->italic = $italic;

        return $this;
    }

    public function isItalic()
    {
        return $this->italic;
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

    public function setUnderline($underline)
    {
        if (!in_array($underline, $this->getUnderlines())) {
            throw new \InvalidArgumentException("The parameter must belong to the Underline parameter list");
        }

        $this->underline = $underline;

        return $this;
    }

    public function getUnderline()
    {
        return $this->underline;
    }

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