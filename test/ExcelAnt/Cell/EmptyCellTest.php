<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\EmptyCell,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font,
    ExcelAnt\Collections\StyleCollection;

class EmptyCellTest extends \PHPUnit_Framework_TestCase
{
    private $emptyCell;

    public function setUp()
    {
        $this->emptyCell = new EmptyCell();
    }

    /**
     * @expectedException \BadMethodCallException
     */
    public function testCannotSetValue()
    {
        $this->emptyCell->setValue('foo');
    }

    /**
     * @expectedException \BadMethodCallException
     */
    public function testCannotGetValue()
    {
        $this->emptyCell->getValue();
    }

    public function testAddAndStyle()
    {
        $styleCollection = new StyleCollection([new Fill(), new Font()]);
        $this->emptyCell->setStyles($styleCollection);

        $this->assertInstanceOf('ExcelAnt\Collections\StyleCollection', $this->emptyCell->getStyles());
    }

    public function testHasStyles()
    {
        $styleCollection = new StyleCollection([new Fill(), new Font()]);

        $this->assertFalse($this->emptyCell->hasStyles());

        $this->emptyCell->setStyles($styleCollection);

        $this->assertTrue($this->emptyCell->hasStyles());
    }
}