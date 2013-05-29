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

    public function setRow();

    public function getRow();

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