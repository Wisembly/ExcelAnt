<?php

namespace ExcelAnt\Adapter\PhpExcel;

use \PHPExcel_Worksheet;
use \PHPExcel_Worksheet_RowDimension;
use \PHPExcel_Worksheet_ColumnDimension;

use ExcelAnt\Adapter\PhpExcel\Workbook\Workbook,
    ExcelAnt\Adapter\PhpExcel\Sheet\Sheet,
    ExcelAnt\Table\Table,
    ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Cell\Cell;

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

        $sheet = $this->createSheet(null, $phpExcelWorksheet)->setTitle('Foo');

        $this->assertEquals('Foo', $sheet->getTitle());
    }

    public function testAddAndGetTable()
    {
        $table = new Table();

        $sheet = $this->createSheet()->addTable($table, new Coordinate(1, 1));

        $this->assertCount(1, $sheet->getTables());
    }

    public function testAddAndGetCell()
    {
        $sheet = $this->createSheet()->addCell(new Cell(), new Coordinate(1, 1));

        $this->assertCount(1, $sheet->getCells());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetRowHeightWithWrongHeightParameter()
    {
        $sheet = $this->createSheet()->setRowHeight('foo', 1);
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetRowHeightWithWrongIdParameter()
    {
        $sheet = $this->createSheet()->setRowHeight(1, 'foo');
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetRowHeightWithWrongParameter()
    {
        $sheet = $this->createSheet()->getRowHeight('foo');
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

        $sheet = $this->createSheet(null, $phpExcelWorksheet)->setRowHeight(3, 1);

        $this->assertEquals(3, $sheet->getRowHeight(1));
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetDefaultRowHeightWithWrongHeightParameter()
    {
        $sheet = $this->createSheet()->setDefaultRowHeight('foo');
    }

    public function testSetAndGetDefaultRowHeight()
    {
        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelRowDimension = $this->getPhpExcelRowDimensionMock();

        $phpExcelRowDimension->expects($this->any())
            ->method('setRowHeight')
            ->will($this->returnValue($phpExcelRowDimension));
        $phpExcelRowDimension->expects($this->any())
            ->method('getRowHeight')
            ->will($this->returnValue(30));

        $phpExcelWorksheet->expects($this->any())
            ->method('getDefaultRowDimension')
            ->will($this->returnValue($phpExcelRowDimension));

        $this->assertEquals(30, $this->createSheet(null, $phpExcelWorksheet)->setDefaultRowHeight(30)->getDefaultRowHeight());
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetColumnWidthWithWrongWidthParameter()
    {
        $sheet = $this->createSheet()->setColumnWidth('foo', 0);
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testSetColumnWidthWithWrongIdParameter()
    {
        $sheet = $this->createSheet()->setColumnWidth(0, 'foo');
    }

    /**
     * @expectedException \InvalidArgumentException
     */
    public function testGetColumnWidthWithWrongParameter()
    {
        $sheet = $this->createSheet()->getColumnWidth('foo');
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
            ->method('getColumnDimensionByColumn')
            ->will($this->returnValue($phpExcelRowDimension));

        $sheet = $this->createSheet(null, $phpExcelWorksheet)->setColumnWidth(3, 0);

        $this->assertEquals(3, $sheet->getColumnWidth(1));
    }

    /**
     * Create a Sheet
     *
     * @param  Mock_Workbook
     * @param  Mock_PHPExcel_Worksheet $phpExcelWorksheet
     *
     * @return Sheet
     */
    public function createSheet($workbook = null, $phpExcelWorksheet = null)
    {
        $workbook = $workbook ?: $this->getWorkbookMock();
        $phpExcelWorksheet = $phpExcelWorksheet ?: $this->getPhpExcelWorksheetMock();

        return new Sheet($workbook, $phpExcelWorksheet);
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
     * Mock ExcelAnt\Adapter\PhpExcel\Workbook\Workbook
     * @return Mock
     */
    private function getWorkbookMock()
    {
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Workbook\Workbook')->disableOriginalConstructor()->getMock();
    }
}