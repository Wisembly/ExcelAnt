<?php

namespace ExcelAnt\Sheet;

interface SheetInterface
{
    public function getRawClass();

    public function setTitle($name);

    public function getTitle();

    public function writeCell();

    public function addTable();

    public function setRowHeight();

    public function setColumnWidth();

    public function importImage();

    public function write();
}