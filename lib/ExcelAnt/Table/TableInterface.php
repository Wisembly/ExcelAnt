<?php

namespace ExcelAnt\Table;

use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Collections\StyleCollection;

interface TableInterface
{
    /**
     * Set labels
     *
     * @param array           $labels The data of labels
     * @param string          $type   Use const to define the type of the labels.
     * @param StyleCollection $styles
     *
     * @return TableInterface
     */
    public function setLabels($labels, $type = self::LABEL_TOP, StyleCollection $styles = null);

    /**
     * Get labels
     *
     * @return array Containing Cell classes
     */
    public function getLabels();

    /**
     * Set a row
     *
     * @param array           $data   The row data
     * @param int             $index  The index if you want insert at a specific row
     * @param StyleCollection $styles
     *
     * @throws InvalidException If index isn't numeric
     *
     * @return TableInterface
     */
    public function setRow($data, $index = null, StyleCollection $styles = null);

    /**
     * Get Row
     *
     * @param  int $index
     *
     * @throws InvalidException If index isn't numeric
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
     *
     * @return TableInterface
     */
    public function removeRow($index);

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

    public function setColumn();

    public function getColumn();

    public function getWidth();

    public function getHeight();
}