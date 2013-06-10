<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase,
    ExcelAnt\Style\Border;

class BorderTest extends StyleTestCase
{
    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetTypeWithWrongParameter($type)
    {
        $border = (new Border())->setType($type);
    }

    public function testSetAndGetType()
    {
        $border = (new Border())->setType(Border::BORDER_DASHDOT);

        $this->assertEquals(Border::BORDER_DASHDOT, $border->getType());
    }

    public function testSetAndGetColor()
    {
        $border = (new Border())->setColor('ff0000');

        $this->assertEquals('ff0000', $border->getColor());
    }
}