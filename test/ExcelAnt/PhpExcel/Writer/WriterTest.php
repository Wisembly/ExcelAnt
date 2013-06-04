<?php

namespace ExcelAnt\PhpExcel\Writer;



use ExcelAnt\PhpExcel\Writer\Writer;

class WriterTest extends \PHPUnit_Framework_TestCase
{
    public function testInstanciateWriter()
    {
        $workbook = $this->getWorkbookMock();
        $writer = new Writer($workbook, $this->getTableWorkerMock());

        $this->assertInstanceOf('ExcelAnt\PhpExcel\Workbook', $writer->getWorkbook());
    }

    public function testSetAndGetWriter()
    {
        $workbook = $this->getWorkbookMock();
        $workbook->expects($this->any())
            ->method('getTitle')
            ->will($this->returnValue('Foo'));

        $writer = new Writer($workbook, $this->getTableWorkerMock());

        $workbook = $this->getWorkbookMock();
        $workbook->expects($this->any())
            ->method('getTitle')
            ->will($this->returnValue('Foo'));

        $writer->setWorkbook($workbook);

        $this->assertInstanceOf('ExcelAnt\PhpExcel\Workbook', $writer->getWorkbook());
        $this->assertEquals('Foo', $writer->getWorkbook()->getTitle());
    }

    public function testWrite()
    {
        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('addSheet')
            ->will($this->returnValue($phpExcel));

        $workbook = $this->getWorkbookMock();
        $workbook->expects($this->once())
            ->method('getRawClass')
            ->will($this->returnValue($phpExcel));

        $sheet = $this->getSheetMock();
        $sheet->expects($this->once())
            ->method('getRawClass')
            ->will($this->returnValue($this->getPhpExcelWorksheetMock()));
        $sheet->expects($this->once())
            ->method('getTables')
            ->will($this->returnValue([$this->getTableMock()]));

        $workbook->expects($this->any())
            ->method('getAllSheets')
            ->will($this->returnValue([$sheet]));

        $tableWorker = $this->getTableWorkerMock();
        $tableWorker->expects($this->once())
            ->method('writeTable')
            ->will($this->returnValue($this->getPhpExcelWorksheetMock()));

        $writer = (new Writer($workbook, $tableWorker))->write('foo');
    }

    /**
     * Mock TableWorker
     * @return Mock
     */
    private function getTableWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Writer\TableWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock Workbook
     * @return Mock
     */
    private function getWorkbookMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Workbook')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock Sheet
     * @return Mock
     */
    private function getSheetMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Sheet')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock Table
     * @return Mock
     */
    private function getTableMock()
    {
        return $this->getMockBuilder('ExcelAnt\Table\Table')->disableOriginalConstructor()->getMock();
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
     * Mock PHPExcel
     *
     * @return Mock_PHPExcel
     */
    private function getPhpExcelMock()
    {
        return $this->getMockBuilder('PHPExcel')->disableOriginalConstructor()->getMock();
    }
}