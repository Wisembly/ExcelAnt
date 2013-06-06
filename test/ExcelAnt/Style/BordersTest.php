<?php

namespace ExcelAnt\Style;

use ExcelAnt\Style\Border;

class BordersTest extends \PHPUnit_Framework_TestCase
{
    public function testSetAndGetTop()
    {
        $borders = (new Borders())->setTop(new Border());

        $this->assertInstanceOf('ExcelAnt\Style\Border', $borders->getTop());
    }

    public function testSetAndGetBottom()
    {
        $borders = (new Borders())->setBottom(new Border());

        $this->assertInstanceOf('ExcelAnt\Style\Border', $borders->getBottom());
    }

    public function testSetAndGetLeft()
    {
        $borders = (new Borders())->setLeft(new Border());

        $this->assertInstanceOf('ExcelAnt\Style\Border', $borders->getLeft());
    }

    public function testSetAndGetRight()
    {
        $borders = (new Borders())->setRight(new Border());

        $this->assertInstanceOf('ExcelAnt\Style\Border', $borders->getRight());
    }

    public function testGetBorders()
    {
        $borders = (new Borders())
            ->setRight(new Border())
            ->setLeft(new Border())
            ->setTop(new Border())
            ->setBottom(new Border());

        $this->assertInstanceOf('ExcelAnt\Style\Border', $borders->getBorders()['top']);
        $this->assertInstanceOf('ExcelAnt\Style\Border', $borders->getBorders()['bottom']);
        $this->assertInstanceOf('ExcelAnt\Style\Border', $borders->getBorders()['left']);
        $this->assertInstanceOf('ExcelAnt\Style\Border', $borders->getBorders()['right']);
    }
}