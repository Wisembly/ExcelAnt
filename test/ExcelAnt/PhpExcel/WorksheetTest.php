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
        $sheet = $this->createSheet($this->worksheet, 'Foo');
        $this->worksheet->addSheet($sheet);

        $this->assertCount(1, $this->worksheet->getAllSheets());
        $this->assertEquals('Foo', $this->worksheet->getSheet($this->worksheet->countSheets() - 1)->getTitle());
    }

    public function testAddSheetWithIndex()
    {
        $sheet = $this->createSheet($this->worksheet, 'Foo');
        $this->worksheet->addSheet($sheet);

        $sheet = $this->createSheet($this->worksheet, 'Bar');
        $this->worksheet->addSheet($sheet, 0);

        $this->assertEquals('Bar', $this->worksheet->getSheet(0)->getTitle());
        $this->assertCount(1, $this->worksheet->getAllSheets());
    }

    /**
     * @dataProvider getDataToInsert
     */
    public function testInsertSheet($worksheet, $sheetCollection, $sheetToInsert, $index)
    {
        // Add sheets
        foreach ($sheetCollection as $sheet) {
            $this->worksheet->addSheet($sheet);
        }

        // Insert
        $this->worksheet->addSheet($sheetToInsert, $index, true);

        // Get new data
        $newSheetCollection = $this->worksheet->getAllSheets();
        $countNewSheetCollection = $this->worksheet->countSheets();

        // Asserts
        $this->assertCount(count($sheetCollection) + 1, $newSheetCollection);

        for ($i=0; $i < $countNewSheetCollection; $i++) {
            if ($i < $index) {
                $this->assertEquals($sheetCollection[$i]->getTitle(), $newSheetCollection[$i]->getTitle());

                continue;
            }

            if ($i === $index) {
                $this->assertEquals($sheetToInsert->getTitle(), $newSheetCollection[$i]->getTitle());

                continue;
            }

            $this->assertEquals($sheetCollection[$i - 1]->getTitle(), $newSheetCollection[$i]->getTitle());
        }
    }

    public function getDataToInsert()
    {
        $worksheet = new Worksheet();

        return [
            [$worksheet, [$this->createSheet($worksheet, 'Foo'), $this->createSheet($worksheet, 'Bar'), $this->createSheet($worksheet, 'Baz')], $this->createSheet($worksheet, 'Insert'), 0],
            [$worksheet, [$this->createSheet($worksheet, 'Foo'), $this->createSheet($worksheet, 'Bar'), $this->createSheet($worksheet, 'Baz')], $this->createSheet($worksheet, 'Insert'), 1],
            [$worksheet, [$this->createSheet($worksheet, 'Foo'), $this->createSheet($worksheet, 'Bar'), $this->createSheet($worksheet, 'Baz')], $this->createSheet($worksheet, 'Insert'), 2],
        ];
    }

    private function createSheet($worksheet, $title = 'Foo')
    {
        return (new Sheet($worksheet))->setTitle($title);
    }
}