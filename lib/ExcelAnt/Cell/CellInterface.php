<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Style\StyleInterface;

interface CellInterface
{
    public function setValue($value);

    public function getValue();

    public function setStyle(StyleInterface $style);

    public function getStyle();
}