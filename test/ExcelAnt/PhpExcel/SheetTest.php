<?php

namespace ExcelAnt\PhpExcel;

use PHPExcel_Worksheet;

use ExcelAnt\PhpExcel\Worksheet;
use ExcelAnt\PhpExcel\Sheet;
use ExcelAnt\Table\Table;
use ExcelAnt\Coordinate\Coordinate;

class SheetTest extends \PHPUnit_Framework_TestCase
{
    public function testRawClass()
    {
        $sheet = $this->createSheet();

        $this->assertInstanceOf("PHPExcel_Worksheet", $sheet->getRawClass());
    }

    public function testSetAndGetTitle()
    {
        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->any())
             ->method('setTitle')
             ->will($this->returnValue($phpExcelWorksheet));
        $phpExcelWorksheet->expects($this->any())
             ->method('getTitle')
             ->will($this->returnValue('Foo'));

        $sheet = $this->createSheet(null, $phpExcelWorksheet);
        $sheet->setTitle('Foo');

        $this->assertEquals('Foo', $sheet->getTitle());
    }

    public function testAddAndGetTable()
    {
        $table = new Table();
        $coordinate = new Coordinate(1, 1);

        $sheet = $this->createSheet();
        $sheet->addTable($table, $coordinate);

        $this->assertCount(1, $sheet->getTables());
    }

    /**
     * Create a Sheet
     * @param  Mock_Worksheet
     * @param  Mock_PHPExcel_Worksheet $phpExcelWorksheet
     * @return Sheet
     */
    public function createSheet($worksheet = null, $phpExcelWorksheet = null)
    {
        $worksheet = $worksheet ?: $this->getWorksheetMock();
        $phpExcelWorksheet = $phpExcelWorksheet ?: $this->getPhpExcelWorksheetMock();

        return new Sheet($worksheet, $phpExcelWorksheet);
    }

    /**
     * Mock PHPExcel_Worksheet
     * @return Mock
     */
    private function getPhpExcelWorksheetMock()
    {
        return $this->getMockBuilder('PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock ExcelAnt\PhpExcel\Worksheet
     * @return Mock
     */
    private function getWorksheetMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Worksheet')->disableOriginalConstructor()->getMock();
    }
}