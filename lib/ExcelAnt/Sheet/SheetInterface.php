<?php

namespace ExcelAnt\Sheet;

use ExcelAnt\Table\TableInterface;
use ExcelAnt\Coordinate\Coordinate;

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

    /**
     * Add a Table
     *
     * @param TableInterface $table
     * @param Coordinate     $coordinate
     *
     * @return SheetInterface
     */
    public function addTable(TableInterface $table, Coordinate $coordinate);

    /**
     * Get tables
     *
     * @return array
     */
    public function getTables();

    public function setRowHeight();

    public function setColumnWidth();

    public function importImage();

    public function write();
}