<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\TableInterface;
use ExcelAnt\Table\LabelInterface;
use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Cell\Cell;
use ExcelAnt\Cell\EmptyCell;
use ExcelAnt\Collections\StyleCollection;

class Table implements TableInterface
{
    private $table = [];
    private $label;
    private $cellCollection = [];

    public function __construct()
    {

    }

    /**
     * {@inheritdoc}
     */
    public function setLabel(LabelInterface $label)
    {
        $this->label = $label;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getLabel()
    {
        return $this->label;
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
            if (false === filter_var($index, FILTER_VALIDATE_INT)) {
                throw new \InvalidArgumentException("Index must be numeric");
            }

            if (array_key_exists($index, $this->table)) {
                $this->cleanRow($index);
                $dataLength = count($data);

                for ($i = 0; $i < $dataLength; $i++) {

                    if (null === $data[$i]) {
                        $cell = new EmptyCell();
                    } else {
                        $cell = new Cell($data[$i]);
                    }

                    if (null !== $styles) {
                        $cell->setStyles($styles);
                    }

                    $this->table[$index][$i] = $cell;
                }

                return $this;
            }

            $this->createNewRow($data, $index, $styles);

            return $this;
        }

        $index = $this->getLastRow();
        $index = null === $index ? 0 : ++$index;

        $this->createNewRow($data, $index, $styles);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getRow($index)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
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

        return current($keys);
    }

    /**
     * {@inheritdoc}
     */
    public function cleanRow($index)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (!isset($this->table[$index])) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        foreach ($this->table[$index] as $key => $value) {
            $this->table[$index][$key] = new EmptyCell();
        }

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function removeRow($index, $reindex = false)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
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

    public function setColumn($data, $index = null, StyleCollection $styles = null)
    {
        if (!is_array($data)) {
            $data = [$data];
        }

        if (null !== $index) {

        } else {
            $index = $this->getLastColumn();
            $index = null === $index ? 0 : ++$index;
        }

        $dataSize = count($data);

        for ($i = 0; $i < $dataSize; $i++) {

            if (null === $data[$i]) {
                $cell = new EmptyCell();
            } else {
                $cell = new Cell($data[$i]);
            }

            $this->table[$i][$index] = $cell;
        }

        return $this;
    }

    public function getColumn($index)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $column = [];

        foreach ($this->table as $row) {
            $column[] = $row[$index];
        }

        if (empty($column)) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        return $column;
    }

    public function getLastColumn()
    {
        if (empty($this->table)) {
            return null;
        }

        $max = 0;

        foreach ($this->table as $row) {
            $size = count($row);

            if ($size > $max) {
                $max = $size;
            }
        }

        return $max - 1;
    }

    public function cleanColumn($index)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (!isset($this->table[$index])) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        foreach ($this->table as $key => $row) {
            if (array_key_exists($index, $row)) {
                $this->table[$key][$index] = new EmptyCell();
            }
        }

        return $this;
    }

    public function removeColumn($index)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (!isset($this->table[$index])) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        $removed = false;

        foreach ($this->table as $key => $row) {
            if (array_key_exists($index, $row)) {
                unset($this->table[$key][$index]);
                $removed = true;
            }
        }

        if (false === $removed) {
            throw new \OutOfBoundsException("Index doesn't exist");
        }

        return $this;
    }

    public function getWidth()
    {

    }

    public function getHeight()
    {

    }

    private function createNewRow($data, $index, $styles = null)
    {
        foreach ($data as $value) {

            if (null === $value) {
                $cell = new EmptyCell();
            } else {
                $cell = new Cell($value);
            }

            if (null !== $styles) {
                $cell->setStyles($styles);
            }

            $this->table[$index][] = $cell;
        }
    }
}