<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleInterface;

class Format implements StyleInterface
{
    /* Data types */
    const TYPE_STRING2    = 'str';
    const TYPE_STRING     = 's';
    const TYPE_FORMULA    = 'f';
    const TYPE_NUMERIC    = 'n';
    const TYPE_NUMERIC_00 = 'nd';
    const TYPE_BOOL       = 'b';
    const TYPE_NULL       = 'null';
    const TYPE_INLINE     = 'inlineStr';
    const TYPE_ERROR      = 'e';
    const TYPE_PERCENT    = 'p';
    const TYPE_PERCENT_00 = 'pd';
    const TYPE_DATETIME   = 'dt';
    const TYPE_DATE       = 'd';

    private $format = self::TYPE_STRING;

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
            self::TYPE_NUMERIC_00,
            self::TYPE_BOOL,
            self::TYPE_NULL,
            self::TYPE_INLINE,
            self::TYPE_ERROR,
            self::TYPE_PERCENT,
            self::TYPE_DATETIME,
            self::TYPE_DATE,
            self::TYPE_PERCENT_00,
        ];
    }
}