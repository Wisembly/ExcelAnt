<?php

namespace ExcelAnt\Coordinate;

class CoordinateTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @expectedException \InvalidArgumentException
     */
    public function testInstanciateWithWrongXAxis()
    {
        new Coordinate('foo', 2);
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testInstanciateWithWrongYAxis()
    {
        new Coordinate(2, 'foo');
    }

    public function testInstanciateAndGetters()
    {
        $coordinate = new Coordinate(1, 2);

        $this->assertEquals(1, $coordinate->getXAxis());
        $this->assertEquals(2, $coordinate->getYAxis());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetXAxisWithWrongParameter()
    {
        $coordinate = new Coordinate(1, 1);
        $coordinate->setXAxis('foo');
    }

    public function testSetAndGetXAxis()
    {
        $coordinate = new Coordinate(1, 1);
        $coordinate->setXAxis(2);

        $this->assertEquals(2, $coordinate->getXAxis());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetYAxisWithWrongParameter()
    {
        $coordinate = new Coordinate(1, 1);
        $coordinate->setYAxis('foo');
    }

    public function testSetAndGetYAxis()
    {
        $coordinate = new Coordinate(1, 1);
        $coordinate->setYAxis(2);

        $this->assertEquals(2, $coordinate->getYAxis());
    }
}