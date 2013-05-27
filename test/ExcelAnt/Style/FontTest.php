<?php

namespace ExcelAnt\Style;

class FontTest extends \PHPUnit_Framework_TestCase
{
    public function testSetAndGetName()
    {
        $font = new Font();
        $font->setName('Arial');

        $this->assertEquals('Arial', $font->getName());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetSizeWithWrongParameter()
    {
        $font = new Font();
        $font->setSize('foo');
    }

    public function testSetAndGetSize()
    {
        $font = new Font();
        $font->setSize(20);

        $this->assertEquals(20, $font->getSize());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetBoldWithWrongParameter()
    {
        $font = new Font();
        $font->setBold('foo');
    }

    public function testEnableAndDisableBold()
    {
        $font = new Font();
        $font->setBold(true);

        $this->assertTrue($font->isBold());

        $font->setBold(false);

        $this->assertFalse($font->isBold());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetItalicWithWrongParameter()
    {
        $font = new Font();
        $font->setItalic('foo');
    }

    public function testEnableAndDisableItalic()
    {
        $font = new Font();
        $font->setItalic(true);

        $this->assertTrue($font->isItalic());

        $font->setItalic(false);

        $this->assertFalse($font->isItalic());
    }

    public function testSetAndGetColor()
    {
        $font = new Font();
        $font->setColor('ff0000');

        $this->assertEquals('ff0000', $font->getColor());
    }

    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetUnderlineWithWongParameter()
    {
        $font = new Font();
        $font->setUnderline('foo');
    }

    public function testSetAndGetUnderline()
    {
        $font = new Font();
        $font->setUnderline(Font::UNDERLINE_DOUBLE);

        $this->assertEquals(Font::UNDERLINE_DOUBLE, $font->getUnderline());
    }

    public function getWrongParameters()
    {
        return [
            ['foo'],
            [''],
            [null],
            ['@&'],
        ];
    }
}