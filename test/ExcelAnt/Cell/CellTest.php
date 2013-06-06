<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\Cell;
use ExcelAnt\Style\StyleInterface;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;
use ExcelAnt\Collections\StyleCollection;

class CellTest extends \PHPUnit_Framework_TestCase
{
    public function testInstanciateWithValue()
    {
        $cell = new Cell('foo');

        $this->assertEquals('foo', $cell->getValue());
    }

    public function testSetAndGetValue()
    {
        $cell = (new Cell())->setValue('foo');

        $this->assertEquals('foo', $cell->getValue());
    }

    public function testAddAndStyle()
    {
        $styleCollection = new StyleCollection([new Fill(), new Font()]);
        $cell = (new Cell())->setStyles($styleCollection);

        $this->assertInstanceOf('ExcelAnt\Collections\StyleCollection', $cell->getStyles());
    }

    public function testHasStyles()
    {
        $styleCollection = new StyleCollection([new Fill(), new Font()]);
        $cell = new Cell();

        $this->assertFalse($cell->hasStyles());

        $cell->setStyles($styleCollection);

        $this->assertTrue($cell->hasStyles());
    }
}