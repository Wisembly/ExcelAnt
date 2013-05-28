<?php

namespace ExcelAnt\Table;

use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Collections\StyleCollection;

interface TableInterface
{
    public function setLabels($labels, $type = self::LABEL_TOP, StyleCollection $styles = null);

    public function getLabels();

    public function setRow();

    public function getRow();

    public function setCell(CellInterface $cell);

    public function getCells();

    public function setColumn();

    public function getColumn();

    public function getWidth();

    public function getHeight();
}