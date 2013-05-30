<?php

namespace ExcelAnt\Sheet;

use ExcelAnt\Table\TableInterface;
use ExcelAnt\Cell\CellInterface;
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

    /**
     * Add a single Cell. Hight priority, override the Tables
     *
     * @param CellInterface $cell
     * @param Coordinate    $coordinate
     *
     * @return SheetInterface
     */
    public function addCell(CellInterface $cell, Coordinate $coordinate);

    public function getCells();

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

    /**
     * Set the height of a row
     *
     * @param int $height
     * @param int $index
     *
     * @throws InvalidException If height isn't numeric
     * @throws InvalidException If index isn't numeric
     *
     * @return SheetInterface
     */
    public function setRowHeight($height, $index);

    /**
     * Get Row height
     *
     * @param  int $index
     *
     * @throws InvalidException If index isn't numeric
     *
     * @return int
     */
    public function getRowHeight($index);

    public function setColumnWidth();

    public function importImage();
}