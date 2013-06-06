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
            ->method('convertStyles');

        $cell = new EmptyCell();

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $this->getPhpExcelWorksheetMock(), new Coordinate(1, 1));
    }

    public function testWriteEmptyCellWithStyles()
    {
        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue(['foo']));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->once())
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));

        $cell = (new EmptyCell())->setStyles(new StyleCollection([new Fill(), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));
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
            ->method('convertStyles')
            ->will($this->returnValue(['foo']));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->once())
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));
        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                $localFormatStorage = $format;
            }));

        $cell = (new Cell())->setStyles(new StyleCollection([new Fill(), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));

        $this->assertEquals(Format::TYPE_STRING, $localFormatStorage);
    }

    public function testWriteCellWithAFormatStyleAndStyle()
    {
        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue(['foo']));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->once())
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));
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

    public function testWriteCellWithAFormatStyleAndNoStyle()
    {
        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue([]));

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->exactly(0))
            ->method('getStyleByColumnAndRow');
        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                $localFormatStorage = $format;
            }));

        $cell = (new Cell())->setStyles(new StyleCollection([(new Format())->setFormat(Format::TYPE_NUMERIC)]));

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
     * Mock PHPExcel_Style
     * @return Mock
     */
    private function getPhpExcelStyleMock()
    {
        return $this->getMockBuilder('PHPExcel_Style')->disableOriginalConstructor()->getMock();
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