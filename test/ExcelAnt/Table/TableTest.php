<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\Table;
use ExcelAnt\Table\Label;
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

    public function testSetAndGetLabel()
    {
        $this->table->setLabel(new Label());

        $this->assertInstanceOf('ExcelAnt\Table\Label', $this->table->getLabel());
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

    public function testGetLastRow()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->setRow(['bob', 'bobby']);
        $this->assertEquals(1, $this->table->getLastRow());
    }

    public function testGetLastRowWithColumn()
    {
        $this->table->setColumn(['foo', 'bar', 'baz']);
        $this->table->setColumn(['bob', 'bobby']);
        $this->assertEquals(2, $this->table->getLastRow());
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

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetColumnWithInvalidIndex()
    {
        $this->table->getColumn('foo');
    }

    /**
     * @expectedException \OutOfBoundsException
     */
    public function testGetColumnWithOutOfBoundsIndex()
    {
        $this->table->getColumn(999999);
    }

    public function testSetAndGetColumnWithDefaultConfiguration()
    {
        $data = ['foo', 'bar', 'baz'];

        $this->table->setColumn($data);
        $column = $this->table->getColumn(0);

        foreach ($column as $key => $cell) {
            $this->assertEquals($data[$key], $cell->getValue());
        }
    }

    public function testSetColumnWithSingleValue()
    {
        $this->table->setColumn('foo');
        $column = $this->table->getColumn(0);

        foreach ($column as $key => $cell) {
            $this->assertEquals('foo', $cell->getValue());
        }
    }

    public function testSetColumnWhenThereAreAlreadyRows()
    {
        $data = ['col1', 'col2', 'col3'];

        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->setRow(['bob', 'bobby']);

        $this->table->setColumn($data);
        $column = $this->table->getColumn(3);

        foreach ($column as $key => $cell) {
            $this->assertEquals($data[$key], $cell->getValue());
        }
    }

    public function testGetLastColumn()
    {
        $this->table->setColumn(['foo', 'bar', 'baz']);
        $this->table->setColumn(['foo', 'bar', 'baz']);

        $this->assertEquals(1, $this->table->getLastColumn());
    }

    public function testGetLastColumnWithRows()
    {
        $this->table->setRow(['foo', 'bar', 'baz']);
        $this->table->setRow(['bob', 'bobby']);

        $this->assertEquals(2, $this->table->getLastColumn());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testCleanColumnWithInvalidIndex()
    {
        $this->table->cleanColumn('foo');
    }

    /**
     * @expectedException \OutOfBoundsException
     */
    public function testCleanColumnWithOutOfBoundsIndex()
    {
        $this->table->cleanColumn(999999);
    }

}