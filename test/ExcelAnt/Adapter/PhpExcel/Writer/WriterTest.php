<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer;

use ExcelAnt\Adapter\PhpExcel\Writer\Writer,
    ExcelAnt\Adapter\PhpExcel\Workbook\Workbook,
    ExcelAnt\Adapter\PhpExcel\Sheet\Sheet,
    ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font,
    ExcelAnt\Table\Table,
    ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Cell\Cell,
    ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface;

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

        $writer = new Writer($this->getPhpExcelWriterInterfaceMock(), $tableWorker, $this->getCellWorkerMock(), $this->getStyleWorkerMock());
        $phpExcel = $writer->convert($workbook);

        $this->assertInstanceOf('PHPExcel', $phpExcel);

        // test the active sheet is the first
        $this->assertEquals(0, $phpExcel->getActiveSheet());
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

        $writer = new Writer($this->getPhpExcelWriterInterfaceMock(), $tableWorker, $cellWorker, $this->getStyleWorkerMock());
        $phpExcel = $writer->convert($workbook);

        $expected = [
            ['foo', 1, 1],
            ['bar', 2, 1],
        ];

        $this->assertEquals($expected, $localCellStorage);
        $this->assertInstanceOf('PHPExcel', $phpExcel);
    }

    public function testApplyTheStylesOfTheWorkbook()
    {
        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcel = $this->getPhpExcelMock();
        $phpExcel->expects($this->once())
            ->method('getDefaultStyle')
            ->will($this->returnValue($phpExcelStyle));

        $workbook = $this->createWorkbook($phpExcel);

        $workbook->addStyles(new StyleCollection([new Fill(), new Font()]));

        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles');

        $writer = new Writer($this->getPhpExcelWriterInterfaceMock(), $this->getTableWorkerMock(), $this->getCellWorkerMock(), $styleWorker);
        $phpExcel = $writer->convert($workbook);

        $this->assertInstanceOf('PHPExcel', $phpExcel);
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
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock PHPExcel_Style
     * @return Mock
     */
    private function getPhpExcelStyleMock()
    {
        return $this->getMockBuilder('PHPExcel_Style')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock TableWorker
     *
     * @return Mock_TableWorker
     */
    public function getTableWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Writer\Worker\TableWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock CellWorker
     *
     * @return Mock_CellWorker
     */
    public function getCellWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Writer\Worker\CellWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock StyleWorker
     *
     * @return Mock_StyleWorker
     */
    public function getStyleWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Writer\Worker\StyleWorker')->disableOriginalConstructor()->getMock();
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