<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\PhpExcel\Writer\Writer;
use ExcelAnt\PhpExcel\Workbook;
use ExcelAnt\PhpExcel\Sheet;
use ExcelAnt\Table\Table;
use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Cell\Cell;
use ExcelAnt\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface;

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

        $writer = new Writer($tableWorker, $this->getCellWorkerMock());
        $writer->write($workbook, $this->getPhpExcelWriterInterfaceMock());
    }

    public function testWriteASingleCell()
    {
        $cellWorker = $this->getCellWorkerMock();
        $cellWorker->expects($this->exactly(2))
            ->method('writeCell')
            ->will($this->returnCallback(function($cell, $phpExcelWorksheet, $coordinate) use (&$localCellStorage) {
                $localCellStorage[] = [$cell->getValue(), $coordinate->getXAxis(), $coordinate->getYAxis()];
            }));

        $tableWorker = $this->getTableWorkerMock();
        $tableWorker->expects($this->exactly(0))
            ->method('writeTable');

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('addSheet');

        $workbook = $this->createWorkbook($phpExcel);
        $sheet = (new Sheet($workbook, $this->getPhpExcelWorksheetMock()))
            ->addCell((new Cell())->setValue('foo'), new Coordinate(1, 1))
            ->addCell((new Cell())->setValue('bar'), new Coordinate(2, 1));
        $workbook->addSheet($sheet);

        $writer = new Writer($tableWorker, $cellWorker);
        $writer->write($workbook, $this->getPhpExcelWriterInterfaceMock());

        $expected = [
            ['foo', 1, 1],
            ['bar', 2, 1],
        ];

        $this->assertEquals($expected, $localCellStorage);
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
     * Mock PhpExcelWriterInterface
     *
     * @return Mock_PhpExcelWriterInterface
     */
    public function getPhpExcelWriterInterfaceMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock TableWorker
     *
     * @return Mock_TableWorker
     */
    public function getTableWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Writer\Worker\TableWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock CellWorker
     *
     * @return Mock_CellWorker
     */
    public function getCellWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Writer\Worker\CellWorker')->disableOriginalConstructor()->getMock();
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