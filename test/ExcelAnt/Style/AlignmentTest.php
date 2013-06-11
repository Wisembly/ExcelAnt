<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase,
    ExcelAnt\Style\Alignment;

class AlignmentTest extends StyleTestCase
{
    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetVerticalWithWrongParameter($param)
    {
        $alignment = (new Alignment())->setVertical($param);
    }

    public function testSetAndGetVertical()
    {
        $alignment = (new Alignment())->setVertical(Alignment::VERTICAL_BOTTOM);

        $this->assertEquals(Alignment::VERTICAL_BOTTOM, $alignment->getVertical());
    }

    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetHorizontalWithWrongParameter($param)
    {
        $alignment = (new Alignment())->setHorizontal($param);
    }

    public function testSetAndGetHorizontal()
    {
        $alignment = (new Alignment())->setHorizontal(Alignment::HORIZONTAL_LEFT);

        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $alignment->getHorizontal());
    }
}