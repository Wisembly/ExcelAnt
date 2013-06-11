<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase,
    ExcelAnt\Style\Font;

class FontTest extends StyleTestCase
{
    public function testSetAndGetName()
    {
        $font = (new Font())->setName('Arial');

        $this->assertEquals('Arial', $font->getName());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetSizeWithWrongParameter()
    {
        $font = (new Font())->setSize('foo');
    }

    public function testSetAndGetSize()
    {
        $font = (new Font())->setSize(20);

        $this->assertEquals(20, $font->getSize());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetBoldWithWrongParameter()
    {
        $font = (new Font())->setBold('foo');
    }

    public function testEnableAndDisableBold()
    {
        $font = (new Font())->setBold(true);

        $this->assertTrue($font->isBold());

        $font->setBold(false);

        $this->assertFalse($font->isBold());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetItalicWithWrongParameter()
    {
        $font = (new Font())->setItalic('foo');
    }

    public function testEnableAndDisableItalic()
    {
        $font = (new Font())->setItalic(true);

        $this->assertTrue($font->isItalic());

        $font->setItalic(false);

        $this->assertFalse($font->isItalic());
    }

    public function testSetAndGetColor()
    {
        $font = (new Font())->setColor('ff0000');

        $this->assertEquals('ff0000', $font->getColor());
    }

    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetUnderlineWithWongParameter()
    {
        $font = (new Font())->setUnderline('foo');
    }

    public function testSetAndGetUnderline()
    {
        $font = (new Font())->setUnderline(Font::UNDERLINE_DOUBLE);

        $this->assertEquals(Font::UNDERLINE_DOUBLE, $font->getUnderline());
    }
}