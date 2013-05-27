<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\StyleTestCase;
use ExcelAnt\Style\Alignment;

class AlignmentTest extends StyleTestCase
{
    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetVerticalWithWrongParameter($param)
    {
        $alignment = new Alignment();
        $alignment->setVertical($param);
    }

    public function testSetAndGetVertical()
    {
        $alignment = new Alignment();
        $alignment->setVertical(Alignment::VERTICAL_BOTTOM);

        $this->assertEquals(Alignment::VERTICAL_BOTTOM, $alignment->getVertical());
    }

    /**
     * @dataProvider getWrongParameters
     * @expectedException \InvalidArgumentException
     */
    public function testSetHorizontalWithWrongParameter($param)
    {
        $alignment = new Alignment();
        $alignment->setHorizontal($param);
    }

    public function testSetAndGetHorizontal()
    {
        $alignment = new Alignment();
        $alignment->setHorizontal(Alignment::HORIZONTAL_LEFT);

        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $alignment->getHorizontal());
    }
}