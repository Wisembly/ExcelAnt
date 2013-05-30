<?php

namespace ExcelAnt\Table;

use ExcelAnt\Table\Label;
use ExcelAnt\Cell\Cell;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;
use ExcelAnt\Collections\StyleCollection;

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
}