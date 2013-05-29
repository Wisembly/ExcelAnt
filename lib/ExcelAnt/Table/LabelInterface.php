<?php

namespace ExcelAnt\Table;

use ExcelAnt\Collections\StyleCollection;

interface LabelInterface
{
    const TOP  = 'top';
    const LEFT = 'left';
    const FULL = 'full';

    public function setType($type);

    public function getType();

    public function getTypes();

    public function setValues($values, StyleCollection $styles = null);

    public function getValues();
}