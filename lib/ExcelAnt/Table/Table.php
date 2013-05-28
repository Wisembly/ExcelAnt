<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\TableInterface;
use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Cell\Cell;
use ExcelAnt\Collections\StyleCollection;

class Table implements TableInterface
{
    const LABEL_TOP = 'top';

    private $labels;
    private $cellCollection;

    public function __construct()
    {

    }

    public function setLabels($labels, $type = self::LABEL_TOP, StyleCollection $styles = null)
    {
        foreach ($labels as $value) {
            $this->labels[] = (new Cell($value))->setStyles($styles);
        }
    }

    public function getLabels()
    {
        return $this->labels;
    }

    public function setRow()
    {

    }

    public function getRow()
    {

    }

    public function setCell(CellInterface $cell)
    {
        $this->cellCollection[] = $cell;

        return $this;
    }

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