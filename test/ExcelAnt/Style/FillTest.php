<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase;
use ExcelAnt\Style\Fill;

class FillTest extends StyleTestCase
{
    public function testInstanciateWithColor()
    {
        new Fill('ff0000');
    }

    public function testSetAndGetColor()
    {
        $fill = new Fill();
        $fill->setColor('ff0000');

        $this->assertEquals('ff0000', $fill->getColor());
    }
}