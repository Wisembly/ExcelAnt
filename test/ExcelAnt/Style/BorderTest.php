<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase;
use ExcelAnt\Style\Border;

class BorderTest extends StyleTestCase
{
    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testInstanciateWithWrongSideParameter($side)
    {
        new Border($side);
    }

    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetSideWithWrongSideParameter($side)
    {
        $border = new Border(Border::SIDE_LEFT);
        $border->setSide($side);
    }

    public function testSetAndGetSide()
    {
        $border = new Border(Border::SIDE_LEFT);
        $border->setSide(Border::SIDE_RIGHT);

        $this->assertEquals(Border::SIDE_RIGHT, $border->getSide());
    }

    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetTypeWithWrongParameter($type)
    {
        $border = new Border(Border::SIDE_LEFT);
        $border->setType($type);
    }

    public function testSetAndGetType()
    {
        $border = new Border(Border::SIDE_LEFT);
        $border->setType(Border::BORDER_DASHDOT);

        $this->assertEquals(Border::BORDER_DASHDOT, $border->getType());
    }

    public function testSetAndGetColor()
    {
        $border = new Border(Border::SIDE_LEFT);
        $border->setColor('ff0000');

        $this->assertEquals('ff0000', $border->getColor());
    }
}