<?php

namespace ExcelAnt\Traits;

use ExcelAnt\Coordinate\Coordinate;

trait Coordinable
{
    protected $coordinate;

    /**
     * {@inheritdoc}
     */
    public function setCoordinate(Coordinate $coordinate)
    {
        $this->coordinate = $coordinate;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getCoordinate()
    {
        return $this->coordinate;
    }
}