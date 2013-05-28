<?php

namespace ExcelAnt\Table;

interface TableInterface
{
    public function setLabel();

    public function getLabel();

    public function setRow();

    public function getRow();

    public function setCell();

    public function getCells();

    public function setColumn();

    public function getColumn();

    public function getWidth();

    public function getHeight();
}