<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\Cell;
use ExcelAnt\Style\StyleInterface;

class CellTest extends \PHPUnit_Framework_TestCase
{
    public function testInstanciateWithValue()
    {
        $cell = new Cell('foo');

        $this->assertEquals('foo', $cell->getValue());
    }

    public function testSetAndGetValue()
    {
        $cell = new Cell();
        $cell->setValue('foo');

        $this->assertEquals('foo', $cell->getValue());
    }

    // public function testAddAndStyle()
    // {
    //     $style = $this->getMockBuilder('ExcelAnt\Style\StyleInterface')->disableOriginalConstructor()->getMock();
    //     $cell = new Cell();
    //     $cell->setStyle($style);

    //     $this->assertInstanceOf('ExcelAnt\Style\StyleInterface', $cell->getStyle());
    // }
}