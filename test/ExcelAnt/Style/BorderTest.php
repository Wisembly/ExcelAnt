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
    public function testSetTypeWithWrongParameter($type)
    {
        $border = (new Border(Border::SIDE_LEFT))->setType($type);
    }

    public function testSetAndGetType()
    {
        $border = (new Border(Border::SIDE_LEFT))->setType(Border::BORDER_DASHDOT);

        $this->assertEquals(Border::BORDER_DASHDOT, $border->getType());
    }

    public function testSetAndGetColor()
    {
        $border = (new Border(Border::SIDE_LEFT))->setColor('ff0000');

        $this->assertEquals('ff0000', $border->getColor());
    }
}