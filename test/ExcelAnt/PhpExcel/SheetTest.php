<?php

namespace ExcelAnt\PhpExcel;

use ExcelAnt\PhpExcel\Worksheet;
use ExcelAnt\PhpExcel\Sheet;

class SheetTest extends \PHPUnit_Framework_TestCase
{
    private $sheet;

    public function setUp()
    {
        $this->sheet = new Sheet(new Worksheet());
    }

    public function testRawClass()
    {
        $this->assertInstanceOf("PHPExcel_Worksheet", $this->sheet->getRawClass());
        $this->assertInstanceOf("PHPExcel", $this->sheet->getRawClass()->getParent());
    }

    public function testSetAndGetTitle()
    {
        $this->sheet->setTitle('Foo');

        $this->assertEquals('Foo', $this->sheet->getTitle());
    }
}