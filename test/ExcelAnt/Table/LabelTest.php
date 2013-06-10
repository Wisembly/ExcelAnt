<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\Label,
    ExcelAnt\Cell\Cell,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font,
    ExcelAnt\Collections\StyleCollection;

class LabelTest extends \PHPUnit_Framework_TestCase
{
    private $label;

    public function setUp()
    {
        $this->label = new Label();
    }

    public function testInstanciateWithType()
    {
        $label = new Label(Label::LEFT);

        $this->assertEquals(Label::LEFT, $label->getType());
    }

    /**
     * @expectedException \OutOfBoundsException
     */
    public function testSetTypeWithWrongParameter()
    {
        $this->label->setType('foo');
    }

    public function testSetAndGetType()
    {
        $this->label->setType(Label::LEFT);

        $this->assertEquals(Label::LEFT, $this->label->getType());
    }

    public function testSetAndGetValuesWithDefaultConfiguration()
    {
        $data = ['Foo', 'Bar', 'Baz'];

        $this->label->setValues($data);
        $values = $this->label->getValues();

        foreach ($values as $key => $cell) {
            $this->assertInstanceOf('ExcelAnt\Cell\Cell', $cell);
            $this->assertEquals($data[$key], $cell->getValue());
        }
    }

    public function testSetValuesWithStyles()
    {
        $data = ['Foo', 'Bar', 'Baz'];
        $styles = new StyleCollection([new Fill(), new Font()]);

        $this->label->setValues($data, $styles);
        $values = $this->label->getValues();

        foreach ($values as $cell) {
            $this->assertInstanceOf('ExcelAnt\Collections\StyleCollection', $cell->getStyles());
        }
    }

    public function testSetValuesWithEmptyValues()
    {
        $data = ['Foo', null, 'Bar'];

        $this->label->setValues($data);
        $values = $this->label->getValues();

        $this->assertEquals('Foo', $values[0]->getValue());
        $this->assertInstanceOf('ExcelAnt\Cell\EmptyCell', $values[1]);
        $this->assertEquals('Bar', $values[2]->getValue());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetValuesForAFullTypeWithWrongData()
    {
        $data = [['Foo', null, 'Bar'], ['foofoo', 'barbar', 'bazbaz'], ['bob', 'bobby']];
        $this->label->setValues($data);
    }

    public function testSetValuesForAFullType()
    {
        $data = [['Foo', null, 'Bar'], ['foofoo', 'barbar', 'bazbaz']];

        $this->label->setValues($data);
        $values = $this->label->getValues();

        $this->assertEquals('Foo', $values[0][0]->getValue());
        $this->assertInstanceOf('ExcelAnt\Cell\EmptyCell', $values[0][1]);
        $this->assertEquals('Bar', $values[0][2]->getValue());
        $this->assertEquals('foofoo', $values[1][0]->getValue());
        $this->assertEquals('barbar', $values[1][1]->getValue());
        $this->assertEquals('bazbaz', $values[1][2]->getValue());
    }

    public function testSetValuesForAFullTypeWithOneArrayAndOneSingleData()
    {
        $data = [['Foo', null, 'Bar'], 'foo'];

        $this->label->setValues($data);
        $values = $this->label->getValues();

        $this->assertEquals('Foo', $values[0][0]->getValue());
        $this->assertInstanceOf('ExcelAnt\Cell\EmptyCell', $values[0][1]);
        $this->assertEquals('Bar', $values[0][2]->getValue());
        $this->assertEquals('foo', $values[1]->getValue());
    }
}