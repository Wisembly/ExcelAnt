<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\LabelInterface;
use ExcelAnt\Cell\Cell;
use ExcelAnt\Cell\EmptyCell;
use ExcelAnt\Collections\StyleCollection;

class Label implements LabelInterface
{
    private $type = self::TOP;
    private $values;

    public function __construct($type = null)
    {
        if (isset($type)) {
            $this->setType($type);
        }
    }

    /**
     * {@inheritdoc}
     */
    public function setType($type)
    {
        if (!in_array($type, $this->getTypes())) {
            throw new \OutOfBoundsException("This type doesn't exist");
        }

        $this->type = $type;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * {@inheritdoc}
     */
    public function getTypes()
    {
        return [
            self::TOP,
            self::LEFT,
            self::FULL,
        ];
    }

    /**
     * {@inheritdoc}
     */
    public function setValues(array $values, StyleCollection $styles = null)
    {
        foreach ($values as $value) {

            if (null === $value) {
                $cell = new EmptyCell();
            } else {
                $cell = new Cell($value);
            }

            if (null !== $styles) {
                $cell->setStyles($styles);
            }

            $this->values[] = $cell;
        }
    }

    /**
     * {@inheritdoc}
     */
    public function getValues()
    {
        return $this->values;
    }
}