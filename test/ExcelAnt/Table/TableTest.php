<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\Table;
use ExcelAnt\Cell\Cell;

class TableTest extends \PHPUnit_Framework_TestCase
{
    public function testSetAndGetCell()
    {
        $table = new Table();
        $table->setCell(new Cell('Foo'));
        $cellCollection = $table->getCells();

        $this->assertCount(1, $cellCollection);
        $this->assertEquals('Foo', $cellCollection[0]->getValue());
    }
}