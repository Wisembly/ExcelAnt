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

    public function setXAxis($xAxis)
    {
        if (false === filter_var($xAxis, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->xAxis = $xAxis;

        return $this;
    }

    public function getXAxis()
    {
        return $this->xAxis;
    }

    public function setYAxis($yAxis)
    {
        if (false === filter_var($yAxis, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->yAxis = $yAxis;
    }

    public function getYAxis()
    {
        return $this->yAxis;
    }
}