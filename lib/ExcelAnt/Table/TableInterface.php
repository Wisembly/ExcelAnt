<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\Label;
use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Collections\StyleCollection;

interface TableInterface
{
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
     * @param array           $data   The row data
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
     * @throws InvalidException If index isn't numeric
     * @throws OutOfBoundsException If index does't exist
     *
     * @return TableInterface
     */
    public function cleanRow($index);

    /**
     * Remove a row. The indexes are re-indexes.
     *
     * @param  int $index
     *
     * @throws InvalidException If index isn't numeric
     * @throws OutOfBoundsException If index does't exist
     *
     * @return TableInterface
     */
    public function removeRow($index, $reindex = false);

    /**
     * Set cell
     *
     * @param CellInterface $cell The cell to add
     */
    public function setCell(CellInterface $cell);

    /**
     * Return the cells
     *
     * @return array Containing Cell classes
     */
    public function getCells();

    public function setColumn($data, $index = null, StyleCollection $styles = null);

    public function getColumn($index);

    public function getLastColumn();

    public function cleanColumn($index);

    public function removeColumn();

    public function getWidth();

    public function getHeight();
}