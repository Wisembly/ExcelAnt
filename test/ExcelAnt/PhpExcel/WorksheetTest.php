<?php

namespace ExcelAnt\PhpExcel;

use ExcelAnt\PhpExcel\Worksheet;
use ExcelAnt\PhpExcel\Sheet;

class WorksheetTest extends \PHPUnit_Framework_TestCase
{
    private $worksheet;

    public function setUp()
    {
        $this->worksheet = new Worksheet();
    }

    public function testRawClass()
    {
        $this->assertInstanceOf("PHPExcel", $this->worksheet->getRawClass());
        $this->assertCount(0, $this->worksheet->getRawClass()->getAllSheets());
    }

    public function testCreateSheet()
    {
        $sheet = $this->worksheet->createSheet();

        $this->assertInstanceOf("ExcelAnt\PhpExcel\Sheet", $sheet);
        $this->assertCount(1, $this->worksheet->getAllSheets());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetSheetWithInvalidArgument()
    {
        $this->worksheet->getSheet("foo");
    }

    /**
     * @expectedException \RuntimeException
     */
    public function testGetSheetWithNonExistentIndex()
    {
        $this->worksheet->getSheet(count($this->worksheet->getAllSheets()) + 1);
    }

    public function testGetSheet()
    {
        $this->worksheet->createSheet();
        $sheet = $this->worksheet->getSheet(0);

        $this->assertInstanceOf("ExcelAnt\PhpExcel\Sheet", $sheet);
    }

    public function testCountSheets()
    {
        $this->worksheet->createSheet();
        $this->worksheet->createSheet();
        $this->worksheet->createSheet();

        $this->assertEquals(3, $this->worksheet->countSheets());
    }

    public function testAddSheetWithoutIndex()
    {
        $sheet = (new Sheet($this->worksheet))->setTitle('Foo');
        $this->worksheet->addSheet($sheet);

        $this->assertCount(1, $this->worksheet->getAllSheets());
        $this->assertEquals('Foo', $this->worksheet->getSheet($this->worksheet->countSheets() - 1)->getTitle());
    }

    public function testAddSheetWithIndex()
    {
        $sheet = (new Sheet($this->worksheet))->setTitle('Foo');
        $this->worksheet->addSheet($sheet);

        $sheet = (new Sheet($this->worksheet))->setTitle('Bar');
        $this->worksheet->addSheet($sheet, 0);

        $this->assertEquals('Bar', $this->worksheet->getSheet(0)->getTitle());
        $this->assertCount(1, $this->worksheet->getAllSheets());
    }

    // public function testInsertSheet()
    // {
    //     $sheet = (new Sheet($this->worksheet))->setTitle('Foo');
    //     $this->worksheet->addSheet($sheet);
    //     $sheet = (new Sheet($this->worksheet))->setTitle('Baz');
    //     $this->worksheet->addSheet($sheet);

    //     $sheet = (new Sheet($this->worksheet))->setTitle('Bar');
    //     $this->worksheet->addSheet($sheet, 1, true);

    //     $this->assertEquals('Foo', $this->worksheet->getSheet(0)->getTitle());
    //     $this->assertEquals('Bar', $this->worksheet->getSheet(1)->getTitle());
    //     $this->assertEquals('Baz', $this->worksheet->getSheet(2)->getTitle());
    //     $this->assertCount(3, $this->worksheet->getAllSheets());
    // }
}