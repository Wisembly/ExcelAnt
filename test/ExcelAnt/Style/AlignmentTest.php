<?php

namespace ExcelAnt\Style;

class AlignmentTest extends \PHPUnit_Framework_TestCase
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
        $alignment->setVertical('bottom');

        $this->assertEquals('bottom', $alignment->getVertical());
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
        $alignment->setHorizontal('left');

        $this->assertEquals('left', $alignment->getHorizontal());
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