<?php

namespace ExcelAnt\Style;

class FillTest extends \PHPUnit_Framework_TestCase
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