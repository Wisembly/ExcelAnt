<?php

namespace ExcelAnt\Style;

class FormatTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetFormatWithWrongParameter($format)
    {
        $format = new Format();
        $format->setFormat($format);
    }

    public function testSetAndGetFormat()
    {
        $format = new Format();
        $format->setFormat(Format::TYPE_NUMERIC);

        $this->assertEquals(Format::TYPE_NUMERIC, $format->getFormat());
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