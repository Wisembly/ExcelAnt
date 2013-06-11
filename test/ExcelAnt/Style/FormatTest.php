<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase,
    ExcelAnt\Style\Format;

class FormatTest extends StyleTestCase
{
    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetFormatWithWrongParameter($format)
    {
        $format = (new Format())->setFormat($format);
    }

    public function testSetAndGetFormat()
    {
        $format = (new Format())->setFormat(Format::TYPE_NUMERIC);

        $this->assertEquals(Format::TYPE_NUMERIC, $format->getFormat());
    }
}