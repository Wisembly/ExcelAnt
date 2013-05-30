<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface;

class Format implements StyleInterface
{
    /* Data types */
    const TYPE_STRING2  = 'str';
    const TYPE_STRING   = 's';
    const TYPE_FORMULA  = 'f';
    const TYPE_NUMERIC  = 'n';
    const TYPE_BOOL     = 'b';
    const TYPE_NULL     = 'null';
    const TYPE_INLINE   = 'inlineStr';
    const TYPE_ERROR    = 'e';

    private $format = self::TYPE_STRING;

    public function __construct()
    {

    }

    /**
     * Set format
     *
     * @param string $format
     *
     * @throws InvalidArgumentException If format doesn't belong the format parameter list
     *
     * @return Format
     */
    public function setFormat($format)
    {
        if (!in_array($format, $this->getFormats())) {
            throw new \InvalidArgumentException("The parameter must belong to the Format parameter list");
        }

        $this->format = $format;

        return $this;
    }

    /**
     * Get format
     *
     * @return string
     */
    public function getFormat()
    {
        return $this->format;
    }

    /**
     * Return the formats parameter list
     *
     * @return array
     */
    public function getFormats()
    {
        return [
            self::TYPE_STRING2,
            self::TYPE_STRING,
            self::TYPE_FORMULA,
            self::TYPE_NUMERIC,
            self::TYPE_BOOL,
            self::TYPE_NULL,
            self::TYPE_INLINE,
            self::TYPE_ERROR,
        ];
    }
}