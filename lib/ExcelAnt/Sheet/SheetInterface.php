<?php

namespace ExcelAnt\Sheet;

interface SheetInterface
{
    public function getRawClass();

    public function setTitle($name);

    public function getTitle();

    public function writeCell();

    public function addTable();

    public function setRowHieght();

    public function setColumnWidth();

    public function importImage();

    public function write();
}