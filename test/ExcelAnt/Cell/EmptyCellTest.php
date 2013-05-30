<?php

namespace ExcelAnt\Cell;

use ExcelAnt\Cell\EmptyCell;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;
use ExcelAnt\Collections\StyleCollection;

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

    /**
     * @expectedException \BadMethodCallException
     */
    public function testCannotSetStyles()
    {
        $this->emptyCell->setStyles(new StyleCollection([new Fill(), new Font()]));
    }

    /**
     * @expectedException \BadMethodCallException
     */
    public function testCannotGetStyles()
    {
        $this->emptyCell->getStyles();
    }
}