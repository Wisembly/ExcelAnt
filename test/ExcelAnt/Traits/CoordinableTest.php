<?php

namespace ExcelAnt\Traits;

use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Table\Table;

class CoordinableTest extends \PHPUnit_Framework_Testcase
{
    public function testSetAndGetCoordinate()
    {
        $table = new Table();
        $table->setCoordinate(new Coordinate(1, 2));
        $coordinate = $table->getCoordinate();

        $this->assertEquals(1, $coordinate->getXAxis());
        $this->assertEquals(2, $coordinate->getYAxis());
    }
}