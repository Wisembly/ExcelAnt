<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\LabelInterface,
    ExcelAnt\Cell\Cell,
    ExcelAnt\Cell\EmptyCell,
    ExcelAnt\Collections\StyleCollection;

class Label implements LabelInterface
{
    private $type = self::TOP;
    private $values;

    /**
     * {@inheritdoc}
     */
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

            // If we have an array here, this is because the client want to create a FULL label
            if (is_array($value)) {

                if (count($values) > 2) {
                    throw new \InvalidArgumentException("If you want to create a full label, you must to pass only an array of two arrays");
                }

                $cells = [];

                foreach ($value as $val) {
                    $cells[] = $this->createCell($val, $styles);
                }

                $this->values[] = $cells;

                continue;
            }

            $this->values[] = $this->createCell($value, $styles);
        }

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getValues()
    {
        return $this->values;
    }

    private function createCell($value = null, StyleCollection $styles = null)
    {
        if (null === $value) {
            $cell = new EmptyCell();
        } else {
            $cell = new Cell($value);
        }

        if (null !== $styles) {
            $cell->setStyles($styles);
        }

        return $cell;
    }
}