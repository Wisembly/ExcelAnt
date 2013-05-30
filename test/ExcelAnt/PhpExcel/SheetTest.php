<?php

namespace ExcelAnt\PhpExcel;

use PHPExcel_Worksheet;
use PHPExcel_Worksheet_RowDimension;
use PHPExcel_Worksheet_ColumnDimension;

use ExcelAnt\PhpExcel\Worksheet;
use ExcelAnt\PhpExcel\Sheet;
use ExcelAnt\Table\Table;
use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Cell\Cell;

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

    public function testAddAndGetCell()
    {
        $sheet = $this->createSheet();
        $sheet->addCell(new Cell(), new Coordinate(1, 1));

        $this->assertCount(1, $sheet->getCells());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetRowHeightWithWrongHeightParameter()
    {
        $sheet = $this->createSheet();
        $sheet->setRowHeight('foo', 1);
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetRowHeightWithWrongIdParameter()
    {
        $sheet = $this->createSheet();
        $sheet->setRowHeight(1, 'foo');
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetRowHeightWithWrongParameter()
    {
        $sheet = $this->createSheet();
        $sheet->getRowHeight('foo');
    }

    public function testSetAndGetRowHeight()
    {
        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelRowDimension = $this->getPhpExcelRowDimensionMock();

        $phpExcelRowDimension->expects($this->any())
            ->method('setRowHeight')
            ->will($this->returnValue($phpExcelRowDimension));
        $phpExcelRowDimension->expects($this->any())
            ->method('getRowHeight')
            ->will($this->returnValue(3));

        $phpExcelWorksheet->expects($this->any())
            ->method('getRowDimension')
            ->will($this->returnValue($phpExcelRowDimension));

        $sheet = $this->createSheet(null, $phpExcelWorksheet);
        $sheet->setRowHeight(3, 1);

        $this->assertEquals(3, $sheet->getRowHeight(1));
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetColumnWidthWithWrongWidthParameter()
    {
        $sheet = $this->createSheet();
        $sheet->setColumnWidth('foo', 1);
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetColumnWidthWithWrongIdParameter()
    {
        $sheet = $this->createSheet();
        $sheet->setColumnWidth(1, 'foo');
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetColumnWidthWithWrongParameter()
    {
        $sheet = $this->createSheet();
        $sheet->getColumnWidth('foo');
    }

    public function testSetAndGetColumnWidth()
    {
        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelRowDimension = $this->getPhpExcelColumnDimensionMock();

        $phpExcelRowDimension->expects($this->any())
            ->method('setWidth')
            ->will($this->returnValue($phpExcelRowDimension));
        $phpExcelRowDimension->expects($this->any())
            ->method('getWidth')
            ->will($this->returnValue(3));

        $phpExcelWorksheet->expects($this->any())
            ->method('getColumnDimension')
            ->will($this->returnValue($phpExcelRowDimension));

        $sheet = $this->createSheet(null, $phpExcelWorksheet);
        $sheet->setColumnWidth(3, 1);

        $this->assertEquals(3, $sheet->getColumnWidth(1));
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
     * Mock PHPExcel_Worksheet_RowDimension
     * @return Mock
     */
    private function getPhpExcelRowDimensionMock()
    {
        return $this->getMockBuilder('PHPExcel_Worksheet_RowDimension')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock PHPExcel_Worksheet_ColumnDimension
     * @return Mock
     */
    private function getPhpExcelColumnDimensionMock()
    {
        return $this->getMockBuilder('PHPExcel_Worksheet_ColumnDimension')->disableOriginalConstructor()->getMock();
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