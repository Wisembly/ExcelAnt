<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\Table;
use ExcelAnt\Cell\Cell;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;
use ExcelAnt\Collections\StyleCollection;

class TableTest extends \PHPUnit_Framework_TestCase
{
    private $table;

    public function setUp()
    {
        $this->table = new Table();
    }

    public function testSetLabelWithDefaultConfiguration()
    {
        $labelsInput = ['Foo', 'Bar', 'Baz'];

        $this->table->setLabels($labelsInput);
        $labels = $this->table->getLabels();

        foreach ($labels as $key => $label) {
            $this->assertInstanceOf('ExcelAnt\Cell\Cell', $label);
            $this->assertEquals($labelsInput[$key], $label->getValue());
        }
    }

    public function testSetLabelWithStyles()
    {
        $labelsInput = ['Foo', 'Bar', 'Baz'];
        $styleCollection = new StyleCollection([new Fill(), new Font()]);

        $this->table->setLabels($labelsInput, null, $styleCollection);
        $labels = $this->table->getLabels();

        foreach ($labels as $cell) {
            $this->assertInstanceOf('ExcelAnt\Collections\StyleCollection', $cell->getStyles());
        }
    }

    public function testSetAndGetCell()
    {
        $this->table->setCell(new Cell('Foo'));
        $cellCollection = $this->table->getCells();

        $this->assertCount(1, $cellCollection);
        $this->assertEquals('Foo', $cellCollection[0]->getValue());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetRowWithInvalidIndex()
    {
        $this->table->getRow('foo');
    }

    /**
     * @expectedException \OutOfBoundsException
     */
    public function testGetRowWithOutOfBoundsIndex()
    {
        $this->table->getRow(999999);
    }

    public function testSetRowWithDefaultConfigurationAndIndirectlyGetRow()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $row = $this->table->getRow(0);

        $this->assertEquals('foo', $row[0]->getValue());
        $this->assertEquals('bar', $row[1]->getValue());
        $this->assertEquals('baz', $row[2]->getValue());
    }

    public function testSetRowWithSingleValue()
    {
        $this->table->setRow('foo');
        $row = $this->table->getRow(0);

        $this->assertEquals('foo', $row[0]->getValue());
    }

    public function testSetRowWhenWhereAreAlreadyRows()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->setRow(['bob', 'bobby']);
        $row = $this->table->getRow(1);

        $this->assertEquals('bob', $row[0]->getValue());
        $this->assertEquals('bobby', $row[1]->getValue());
    }

    public function testLastRow()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->setRow(['bob', 'bobby']);
        $this->assertEquals(1, $this->table->getLastRow());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testCleanRowWithInvalidIndex()
    {
        $this->table->cleanRow('foo');
    }

    /**
     * @expectedException \OutOfBoundsException
     */
    public function testCleanRowWithOutOfBoundsIndex()
    {
        $this->table->cleanRow(999999);
    }

    public function testCleanRow()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->cleanRow(0);

        $this->assertEmpty($this->table->getRow(0));
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetRowWithInvalidIndex()
    {
        $this->table->setRow(['foo', 'bar', 'baz'], 'foo');
    }

    /**
     * @expectedException \OutOfBoundsException
     */
    public function testSetRowWithOutOfBoundsIndex()
    {
        $this->table->setRow(['foo', 'bar', 'baz'], 999999);
    }

    public function testSetRowWithIndex()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->setRow(['bob', 'bobby'], 0);
        $row = $this->table->getRow(0);

        $this->assertEquals('bob', $row[0]->getValue());
        $this->assertEquals('bobby', $row[1]->getValue());
        $this->assertCount(2, $row);
    }

    public function testSetRowWithStyles()
    {
        $styles = new StyleCollection([new Fill(), new Font()]);
        $this->table->setRow(['foo', 'bar', 'baz'], null, $styles);
        $row = $this->table->getRow(0);

        foreach ($row as $cell) {
            $this->assertInstanceOf('ExcelAnt\Collections\StyleCollection', $cell->getStyles());
        }
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testRemoveRowWithInvalidIndex()
    {
        $this->table->removeRow('foo');
    }

    /**
     * @expectedException \OutOfBoundsException
     */
    public function testRemoveRowWithOutOfBoundsIndex()
    {
        $this->table->removeRow(999999);
    }

    public function testRemoveRow()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->setRow(['bob', 'bobby']);
        $this->table->removeRow(0);

        $this->assertCount(1, $this->table->getTable());

        $row = $this->table->getRow(1);
        $this->assertEquals('bob', $row[0]->getValue());
        $this->assertEquals('bobby', $row[1]->getValue());
    }

    public function testRemoveRowWithReIndex()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->setRow(['bob', 'bobby']);
        $this->table->removeRow(0, true);

        $this->assertCount(1, $this->table->getTable());

        $row = $this->table->getRow(0);
        $this->assertEquals('bob', $row[0]->getValue());
        $this->assertEquals('bobby', $row[1]->getValue());
    }

}