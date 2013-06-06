<?php

namespace ExcelAnt\Table;

use ExcelAnt\Collections\StyleCollection;

interface LabelInterface
{
    const TOP  = 'top';
    const LEFT = 'left';
    const FULL = 'full';

    /**
     * @param string $type The label type
     */
    public function __construct($type = null);

    /**
     * Set type
     *
     * @param string $type
     *
     * @return LabelInterface
     */
    public function setType($type);

    /**
     * Get type
     *
     * @return string
     */
    public function getType();

    /**
     * Return all allowed types
     *
     * @return array
     */
    public function getTypes();

    /**
     * Set Value
     *
     * @param array           $values
     * @param StyleCollection $styles
     */
    public function setValues(array $values, StyleCollection $styles = null);

    /**
     * Get values
     *
     * @return array
     */
    public function getValues();
}