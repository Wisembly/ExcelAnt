<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\LabelInterface;
use ExcelAnt\Cell\Cell;
use ExcelAnt\Collections\StyleCollection;

class Label implements LabelInterface
{
    private $type;
    private $values;

    public function __construct($type = null)
    {
        if (isset($type)) {
            $this->setType($type);
        }
    }

    public function setType($type)
    {
        if (!in_array($type, $this->getTypes())) {
            throw new \OutOfBoundsException("This type doesn't exist");
        }

        $this->type = $type;

        return $this;
    }

    public function getType()
    {
        return $this->type;
    }

    public function getTypes()
    {
        return [
            self::TOP,
            self::LEFT,
            self::FULL,
        ];
    }

    public function setValues(array $values, StyleCollection $styles = null)
    {
        foreach ($values as $value) {
            $cell = new Cell($value);

            if (null !== $styles) {
                $cell->setStyles($styles);
            }

            $this->values[] = $cell;
        }
    }

    public function getValues()
    {
        return $this->values;
    }
}