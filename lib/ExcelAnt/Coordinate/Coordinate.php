<?php

namespace ExcelAnt\Coordinate;

class Coordinate
{
    private $xAxis;
    private $yAxis;

    public function __construct($xAxis, $yAxis)
    {
        if (false === filter_var($xAxis, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (false === filter_var($yAxis, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->xAxis = $xAxis;
        $this->yAxis = $yAxis;
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
}