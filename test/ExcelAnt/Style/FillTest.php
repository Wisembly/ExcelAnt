<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase;
use ExcelAnt\Style\Fill;

class FillTest extends StyleTestCase
{
    public function testSetAndGetColor()
    {
        $fill = (new Fill())->setColor('ff0000');

        $this->assertEquals('ff0000', $fill->getColor());
    }
}