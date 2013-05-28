<?php

namespace ExcelAnt\Sheet;

interface SheetInterface
{
    public function getRawClass();

    /**
     * Set title
     *
     * @param string $name The title
     *
     * @return SheetInterface
     */
    public function setTitle($name);

    /**
     * Get title
     *
     * @return string
     */
    public function getTitle();

    public function writeCell();

    public function addTable();

    public function setRowHeight();

    public function setColumnWidth();

    public function importImage();

    public function write();
}