<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\TableInterface;
use ExcelAnt\Cell\CellInterface;

class Table implements TableInterface
{
    private $cellCollection;

    public function __construct()
    {

    }

    public function setLabel()
    {

    }

    public function getLabel()
    {

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