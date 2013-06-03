<?php

namespace ExcelAnt\Coordinate;

class Coordinate
{
    private $xAxis;
    private $yAxis;
    private $originalXAxis;
    private $originalYAxis;

    public function __construct($xAxis, $yAxis)
    {
        if (false === filter_var($xAxis, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (false === filter_var($yAxis, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->xAxis = $this->originalXAxis= $xAxis;
        $this->yAxis = $this->originalYAxis = $yAxis;
    }

    /**
     * Get the original X Axis
     *
     * @return int
     */
    public function getOriginalXAxis()
    {
        return $this->originalXAxis;
    }

    /**
     * Get the original Y Axis
     *
     * @return int
     */
    public function getOriginalYAxis()
    {
        return $this->originalYAxis;
    }

    /**
     * Reset the xAxis from its original value
     *
     * @return Coordinate
     */
    public function resetXAxis()
    {
        $this->xAxis = $this->originalXAxis;

        return $this;
    }

    /**
     * Reset the yAxis from its original value
     *
     * @return Coordinate
     */
    public function resetYAxis()
    {
        $this->yAxis = $this->originalYAxis;

        return $this;
    }

    /**
     * Set X Axis
     *
     * @param int $xAxis
     *
     * @throws InvalidException If xAxis isn't numeric
     *
     * @return Coordinate
     */
    public function setXAxis($xAxis)
    {
        if (false === filter_var($xAxis, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->xAxis = $xAxis;

        return $this;
    }

    /**
     * Get xAxis
     *
     * @return int
     */
    public function getXAxis()
    {
        return $this->xAxis;
    }

    /**
     * Move the X axis by one or more if you give an index
     *
     * @param  integer $index Numeric index of the move
     *
     * @throws InvalidException If index isn't numeric
     *
     * @return Coordinate
     */
    public function nextXAxis($index = null)
    {
        if (isset($index) && false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $index = null === $index ? 1 : $index;
        $this->xAxis = $this->xAxis + $index;

        return $this;
    }

    /**
     * Set Y Axis
     *
     * @param int $yAxis
     *
     * @throws InvalidException If yAxis isn't numeric
     *
     * @return Coordinate
     */
    public function setYAxis($yAxis)
    {
        if (false === filter_var($yAxis, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->yAxis = $yAxis;
    }

    /**
     * Get yAxis
     *
     * @return int
     */
    public function getYAxis()
    {
        return $this->yAxis;
    }

    /**
     * Move the Y axis by one or more if you give an index
     *
     * @param  integer $index Numeric index of the move
     *
     * @throws InvalidException If index isn't numeric
     *
     * @return Coordinate
     */
    public function nextYAxis($index = null)
    {
        if (isset($index) && false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $index = null === $index ? 1 : $index;
        $this->yAxis = $this->yAxis + $index;

        return $this;
    }
}