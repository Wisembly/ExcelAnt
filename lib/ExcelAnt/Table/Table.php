<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\TableInterface;
use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Cell\Cell;
use ExcelAnt\Collections\StyleCollection;

class Table implements TableInterface
{
    private $table = [];
    private $labels = [];
    private $cellCollection = [];

    public function __construct()
    {

    }

    /**
     * {@inheritdoc}
     */
    public function setLabels($labels, $type = self::LABEL_TOP, StyleCollection $styles = null)
    {
        foreach ($labels as $value) {
            $this->labels[] = (new Cell($value))->setStyles($styles);
        }
    }

    /**
     * {@inheritdoc}
     */
    public function getLabels()
    {
        return $this->labels;
    }

    /**
     * {@inheritdoc}
     */
    public function getTable()
    {
        return $this->table;
    }

    /**
     * {@inheritdoc}
     */
    public function setRow($data, $index = null, StyleCollection $styles = null)
    {
        if (!is_array($data)) {
            $data = [$data];
        }

        if (null !== $index) {
            if (!is_numeric($index)) {
                throw new \InvalidArgumentException("Index must be numeric");
            }

            $this->cleanRow($index);
        } else {
            $index = $this->getLastRow();
            $index = null === $index ? 0 : ++$index;
        }

        foreach ($data as $value) {
            $cell = new Cell($value);

            if (null !== $styles) {
                $cell->setStyles($styles);
            }

            $this->table[$index][] = $cell;
        }

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getRow($index)
    {
        if (!is_numeric($index)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (!isset($this->table[$index])) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        return $this->table[$index];
    }

    /**
     * {@inheritdoc}
     */
    public function getLastRow()
    {
        $keys = array_keys($this->table);
        end($keys);

        return key($keys);
    }

    /**
     * {@inheritdoc}
     */
    public function cleanRow($index)
    {
        if (!is_numeric($index)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (!isset($this->table[$index])) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        $this->table[$index] = [];

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function removeRow($index, $reindex = false)
    {
        if (!is_numeric($index)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (!isset($this->table[$index])) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        unset($this->table[$index]);

        if (true === $reindex) {
            $this->table = array_values($this->table);
        }

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setCell(CellInterface $cell)
    {
        $this->cellCollection[] = $cell;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getCells()
    {
        return $this->cellCollection;
    }

    public function setColumn()
    {

    }

    public function getColumn()
    {

    }

    public function getWidth()
    {

    }

    public function getHeight()
    {

    }
}