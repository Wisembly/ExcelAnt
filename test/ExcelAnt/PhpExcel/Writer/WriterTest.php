<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\PhpExcel\Writer\Writer;
use ExcelAnt\PhpExcel\Workbook;
use ExcelAnt\PhpExcel\Sheet;
use ExcelAnt\Table\Table;
use ExcelAnt\Coordinate\Coordinate;

class WriterTest extends \PHPUnit_Framework_TestCase
{
    public function testWriteJustATable()
    {
        $tableWorker = $this->getTableWorkerMock();
        $tableWorker->expects($this->once())
            ->method('writeTable');

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('addSheet');

        $workbook = $this->createWorkbook($phpExcel);
        $sheet = (new Sheet($workbook, $this->getPhpExcelWorksheetMock()))
            ->addTable(new Table(), new Coordinate(1, 1));
        $workbook->addSheet($sheet);

        $writer = new Writer($workbook, $tableWorker);
        $writer->write('foo');
    }

    /**
     * Create a Workbook
     *
     * @param  MockPhpExcel $phpExcel
     *
     * @return Workbook
     */
    public function createWorkbook($phpExcel = null)
    {
        $phpExcel = $phpExcel ?: $this->getPhpExcelMock();

        return new Workbook($phpExcel);
    }

    /**
     * Mock TableWorker
     *
     * @return Mock_TableWorker
     */
    public function getTableWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Writer\TableWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock PHPExcel_Worksheet
     *
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