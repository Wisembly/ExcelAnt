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

    public function testGetOriginalAxis()
    {
        $coordinate = new Coordinate(1, 2);

        $this->assertEquals(1, $coordinate->getOriginalXAxis());
        $this->assertEquals(2, $coordinate->getOriginalYAxis());
    }

    public function testSetOriginalAxis()
    {
        $coordinate = (new Coordinate(1, 2))
            ->setOriginalXAxis(5)
            ->setOriginalYAxis(4);

        $this->assertEquals(5, $coordinate->getOriginalXAxis());
        $this->assertEquals(4, $coordinate->getOriginalYAxis());
    }

    public function testResetXAxis()
    {
        $coordinate = (new Coordinate(1, 2))->setXAxis(5);

        $this->assertEquals(5, $coordinate->getXAxis());
        $coordinate->resetXAxis();
        $this->assertEquals(1, $coordinate->getXAxis());
    }

    public function testResetYAxis()
    {
        $coordinate = (new Coordinate(1, 2))->setYAxis(5);

        $this->assertEquals(5, $coordinate->getYAxis());
        $coordinate->resetYAxis();
        $this->assertEquals(2, $coordinate->getYAxis());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetXAxisWithWrongParameter()
    {
        $coordinate = (new Coordinate(1, 1))->setXAxis('foo');
    }

    public function testSetAndGetXAxis()
    {
        $coordinate = (new Coordinate(1, 1))->setXAxis(2);

        $this->assertEquals(2, $coordinate->getXAxis());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetYAxisWithWrongParameter()
    {
        $coordinate = (new Coordinate(1, 1))->setYAxis('foo');
    }

    public function testSetAndGetYAxis()
    {
        $coordinate = (new Coordinate(1, 1))->setYAxis(2);

        $this->assertEquals(2, $coordinate->getYAxis());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testNexXAxisWithWrongParameter()
    {
        $coordinate = (new Coordinate(1, 1))->nextXAxis('foo');
    }

    public function testNextXAxisWithoutIndex()
    {
        $coordinate = (new Coordinate(1, 1))->nextXAxis();

        $this->assertEquals(2, $coordinate->getXAxis());
    }

    public function testNextXAxisWithIndex()
    {
        $coordinate = (new Coordinate(1, 1))->nextXAxis(3);

        $this->assertEquals(4, $coordinate->getXAxis());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testNexYAxisWithWrongParameter()
    {
        $coordinate = (new Coordinate(1, 1))->nextYAxis('foo');
    }

    public function testNextYAxisWithoutIndex()
    {
        $coordinate = (new Coordinate(1, 1))->nextYAxis();

        $this->assertEquals(2, $coordinate->getYAxis());
    }

    public function testNextYAxisWithIndex()
    {
        $coordinate = (new Coordinate(1, 1))->nextYAxis(3);

        $this->assertEquals(4, $coordinate->getYAxis());
    }
}