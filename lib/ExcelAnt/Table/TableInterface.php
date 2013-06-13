<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\Label,
    ExcelAnt\Cell\CellInterface,
    ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Coordinate\Coordinate;

interface TableInterface
{
    /**
     * Set Coordinate
     *
     * @param Coordinate $coordinate
     *
     * @return TableInterface
     */
    public function setCoordinate(Coordinate $coordinate);

    /**
     * Get Coordinate
     *
     * @return Coordinate
     */
    public function getCoordinate();

    /**
     * Set label
     *
     * @param LabelInterface $label
     *
     * @return TableInterface
     */
    public function setLabel(LabelInterface $label);

    /**
     * Get label
     *
     * @return Label
     */
    public function getLabel();

    /**
     * Set a row
     *
     * @param mixed           $data   The row data
     * @param int             $index  The index if you want insert at a specific row
     * @param StyleCollection $styles
     *
     * @throws InvalidException If index isn't numeric
     * @throws OutOfBoundsException If index does't exist
     *
     * @return TableInterface
     */
    public function setRow($data, $index = null, StyleCollection $styles = null);

    /**
     * Get Row
     *
     * @param  int $index
     *
     * @throws InvalidException     If index isn't numeric
     * @throws OutOfBoundsException If index does't exist
     *
     * @return array
     */
    public function getRow($index);

    /**
     * Get the last row
     *
     * @return int
     */
    public function getLastRow();

    /**
     * Clean a single row. The index already exist
     *
     * @param  int $index
     *
     * @throws InvalidException     If index isn't numeric
     * @throws OutOfBoundsException If index does't exist
     *
     * @return TableInterface
     */
    public function cleanRow($index);

    /**
     * Remove a row.
     *
     * @param  int     $index
     * @param  boolean $reindex If you want reindex the table
     *
     * @throws InvalidException     If index isn't numeric
     * @throws OutOfBoundsException If index does't exist
     *
     * @return TableInterface
     */
    public function removeRow($index, $reindex = false);

    /**
     * Set column
     *
     * @param mixed           $data   The column data
     * @param int             $index  The index if you want insert at a specific row
     * @param StyleCollection $styles
     */
    public function setColumn($data, $index = null, StyleCollection $styles = null);

    /**
     * Get column
     *
     * @param  int $index The index
     *
     * @return array
     */
    public function getColumn($index);

    /**
     * Get last column
     *
     * @return int
     */
    public function getLastColumn();

    /**
     * Clean a single column. The index already exist
     *
     * @param  int $index
     *
     * @throws InvalidException     If index isn't numeric
     * @throws OutOfBoundsException If index does't exist
     *
     * @return TableInterface
     */
    public function cleanColumn($index);

    /**
     * Remove a column.
     *
     * @param  int $index
     *
     * @throws InvalidException     If index isn't numeric
     * @throws OutOfBoundsException If index does't exist
     *
     * @return TableInterface
     */
    public function removeColumn($index);

    /**
     * Return the width of the table
     *
     * @return int Readable and human value
     */
    public function getWidth();

    /**
     * Return the height of the table
     *
     * @return int Readable and human value
     */
    public function getHeight();

    /**
     * Apply a StyleCollection on all cells of a column
     *
     * @param int             $index  The numeric index of the column
     * @param StyleCollection $styles
     *
     * @throws InvalidException     If index isn't numeric
     *
     * @return TableInterface
     */
    public function applyStylesOnColumn($index, StyleCollection $styles = null);

    /**
     * Apply a StyleCollection on all cells of a row
     *
     * @param int             $index  The numeric index of the row
     * @param StyleCollection $styles
     *
     * @throws InvalidException     If index isn't numeric
     *
     * @return TableInterface
     */
    public function applyStylesOnRow($index, StyleCollection $styles = null);
}