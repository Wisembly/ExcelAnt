<?php

namespace ExcelAnt\Sheet;

use ExcelAnt\Table\TableInterface,
    ExcelAnt\Cell\CellInterface,
    ExcelAnt\Coordinate\Coordinate;

interface SheetInterface
{
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

    /**
     * Get cells
     *
     * @return array
     */
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

    /**
     * Set default row height of the sheet
     *
     * @param int $height
     *
     * @throws InvalidException If height isn't numeric
     *
     * @return SheetInterface
     */
    public function setDefaultRowHeight($height);

    /**
     * Get default row height of the sheet
     *
     * @return int
     */
    public function getDefaultRowHeight();

    /**
     * Set the width of a column
     *
     * @param int $width
     * @param int $index
     *
     * @throws InvalidException If width isn't numeric
     * @throws InvalidException If index isn't numeric
     *
     * @return SheetInterface
     */
    public function setColumnWidth($width, $index);

    /**
     * Get Row width
     *
     * @param  int $index
     *
     * @throws InvalidException If index isn't numeric
     *
     * @return int
     */
    public function getColumnWidth($index);

    public function importImage();
}