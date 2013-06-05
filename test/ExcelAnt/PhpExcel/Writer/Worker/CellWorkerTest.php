<?php

namespace ExcelAnt\PhpExcel\Writer\Worker;

use ExcelAnt\PhpExcel\Writer\Worker\CellWorker;
use ExcelAnt\Cell\Cell;
use ExcelAnt\Cell\EmptyCell;
use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Collections\StyleCollection;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;
use ExcelAnt\Style\Format;

class CellWorkerTest extends \PHPUnit_Framework_TestCase
{
    public function testWriteEmptyCell()
    {
        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->exactly(0))
            ->method('applyStyles');

        $cell = new EmptyCell();

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $this->getPhpExcelWorksheetMock(), new Coordinate(1, 1));
    }

    public function testWriteEmptyCellWithStyles()
    {
        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('applyStyles');

        $cell = (new EmptyCell())->setStyles(new StyleCollection([new Fill(), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $this->getPhpExcelWorksheetMock(), new Coordinate(1, 1));
    }

    public function testWriteCell()
    {
        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow');

        $cellWorker = new CellWorker($this->getStyleWorkerMock());
        $cellWorker->writeCell(new Cell(), $phpExcelWorksheet, new Coordinate(1, 1));
    }

    public function testWriteCellWithStyleButNotAFormatStyle()
    {
        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('applyStyles');

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                $localFormatStorage = $format;
            }));

        $cell = (new Cell())->setStyles(new StyleCollection([new Fill(), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));

        $this->assertNull($localFormatStorage);
    }

    public function testWriteCellWithAFormatStyle()
    {
        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('applyStyles');

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                $localFormatStorage = $format;
            }));

        $cell = (new Cell())->setStyles(new StyleCollection([(new Format())->setFormat(Format::TYPE_NUMERIC), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));

        $this->assertEquals('n', $localFormatStorage);
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
     * Mock StyleWorker
     *
     * @return Mock
     */
    private function getStyleWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Writer\Worker\StyleWorker')->disableOriginalConstructor()->getMock();
    }
}