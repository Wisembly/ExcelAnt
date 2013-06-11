<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase,
    ExcelAnt\Style\Fill;

class FillTest extends StyleTestCase
{
    public function testSetAndGetColor()
    {
        $fill = (new Fill())->setColor('ff0000');

        $this->assertEquals('ff0000', $fill->getColor());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetSizeWithWrongParameter()
    {
        $fill = (new Fill())->setType('foo');
    }

    public function testSetAndGetType()
    {
        $fill = (new Fill())->setType(Fill::FILL_SOLID);

        $this->assertEquals(Fill::FILL_SOLID, $fill->getType());
    }
}