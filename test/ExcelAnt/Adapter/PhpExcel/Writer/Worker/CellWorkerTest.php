<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\Worker;

use PHPExcel_Style_NumberFormat;

use ExcelAnt\Adapter\PhpExcel\Writer\Worker\CellWorker,
    ExcelAnt\Cell\Cell,
    ExcelAnt\Cell\EmptyCell,
    ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font,
    ExcelAnt\Style\Format;

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

    public function testWriteCellWithAFormatStyleNumericAndStyle()
    {
        $style = null;

        $phpExcelStyleNumberformat = $this->getPhpExcelStyleNumberFormatMock();
        $phpExcelStyleNumberformat->expects($this->once())
            ->method('setFormatCode')
            ->will($this->returnCallback(function($styleFormat) use (&$style) {
                $style = $styleFormat;
            }));

        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue(['foo']));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcelStyle->expects($this->once())
            ->method('getNumberFormat')
            ->will($this->returnValue($phpExcelStyleNumberformat));

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->exactly(2))
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
        $this->assertEquals(PHPExcel_Style_NumberFormat::FORMAT_NUMBER, $style);
    }

    public function testWriteCellWithAFormatStylePercentageAndStyle()
    {
        $style = null;

        $phpExcelStyleNumberformat = $this->getPhpExcelStyleNumberFormatMock();
        $phpExcelStyleNumberformat->expects($this->once())
            ->method('setFormatCode')
            ->will($this->returnCallback(function($styleFormat) use (&$style) {
                $style = $styleFormat;
            }));

        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue(['foo']));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcelStyle->expects($this->once())
            ->method('getNumberFormat')
            ->will($this->returnValue($phpExcelStyleNumberformat));

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->exactly(2))
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));

        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                $localFormatStorage = $format;
            }));

        $cell = (new Cell())->setStyles(new StyleCollection([(new Format())->setFormat(Format::TYPE_PERCENT), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));

        $this->assertEquals('n', $localFormatStorage);
        $this->assertEquals(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE, $style);
    }

    public function testWriteCellWithAFormatStylePercentage_00AndStyle()
    {
        $style = null;

        $phpExcelStyleNumberformat = $this->getPhpExcelStyleNumberFormatMock();
        $phpExcelStyleNumberformat->expects($this->once())
            ->method('setFormatCode')
            ->will($this->returnCallback(function($styleFormat) use (&$style) {
                $style = $styleFormat;
            }));

        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue(['foo']));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcelStyle->expects($this->once())
            ->method('getNumberFormat')
            ->will($this->returnValue($phpExcelStyleNumberformat));

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->exactly(2))
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));

        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                $localFormatStorage = $format;
            }));

        $cell = (new Cell())->setStyles(new StyleCollection([(new Format())->setFormat(Format::TYPE_PERCENT_00), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));

        $this->assertEquals('n', $localFormatStorage);
        $this->assertEquals(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00, $style);
    }

    public function testWriteCellWithAFormatStyleNumeric_00AndStyle()
    {
        $style = null;

        $phpExcelStyleNumberformat = $this->getPhpExcelStyleNumberFormatMock();
        $phpExcelStyleNumberformat->expects($this->once())
            ->method('setFormatCode')
            ->will($this->returnCallback(function($styleFormat) use (&$style) {
                        $style = $styleFormat;
                    }));

        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue(['foo']));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcelStyle->expects($this->once())
            ->method('getNumberFormat')
            ->will($this->returnValue($phpExcelStyleNumberformat));

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->exactly(2))
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));

        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                        $localFormatStorage = $format;
                    }));

        $cell = (new Cell())->setStyles(new StyleCollection([(new Format())->setFormat(Format::TYPE_NUMERIC_00), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));

        $this->assertEquals('n', $localFormatStorage);
        $this->assertEquals(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00, $style);
    }

    public function testWriteCellWithAFormatStyleDatetimeAndStyle()
    {
        $style = null;

        $phpExcelStyleNumberformat = $this->getPhpExcelStyleNumberFormatMock();
        $phpExcelStyleNumberformat->expects($this->once())
            ->method('setFormatCode')
            ->will($this->returnCallback(function($styleFormat) use (&$style) {
                $style = $styleFormat;
            }));

        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue(['foo']));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('applyFromArray');

        $phpExcelStyle->expects($this->once())
            ->method('getNumberFormat')
            ->will($this->returnValue($phpExcelStyleNumberformat));

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->exactly(2))
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));

        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                $localFormatStorage = $format;
            }));

        $cell = (new Cell())->setStyles(new StyleCollection([(new Format())->setFormat(Format::TYPE_DATETIME), new Font()]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));

        $this->assertEquals('s', $localFormatStorage);
        $this->assertEquals(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME, $style);
    }

    public function testWriteCellWithAFormatStyleAndNoStyle()
    {
        $style = null;

        $phpExcelStyleNumberformat = $this->getPhpExcelStyleNumberFormatMock();
        $phpExcelStyleNumberformat->expects($this->once())
            ->method('setFormatCode')
            ->will($this->returnCallback(function($styleFormat) use (&$style) {
                $style = $styleFormat;
            }));

        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->once())
            ->method('convertStyles')
            ->will($this->returnValue([]));

        $phpExcelStyle = $this->getPhpExcelStyleMock();
        $phpExcelStyle->expects($this->once())
            ->method('getNumberFormat')
            ->will($this->returnValue($phpExcelStyleNumberformat));

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->once())
            ->method('getStyleByColumnAndRow')
            ->will($this->returnValue($phpExcelStyle));

        $phpExcelWorksheet->expects($this->once())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($xAxis, $yAxis, $value, $format) use (&$localFormatStorage) {
                $localFormatStorage = $format;
            }));

        $cell = (new Cell())->setStyles(new StyleCollection([(new Format())->setFormat(Format::TYPE_NUMERIC)]));

        $cellWorker = new CellWorker($styleWorker);
        $cellWorker->writeCell($cell, $phpExcelWorksheet, new Coordinate(1, 1));

        $this->assertEquals('n', $localFormatStorage);
        $this->assertEquals(PHPExcel_Style_NumberFormat::FORMAT_NUMBER, $style);
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
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Writer\Worker\StyleWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock PHPExcel_Style_NumberFormat
     * @return Mock
     */
    private function getPhpExcelStyleNumberFormatMock()
    {
        return $this->getMockBuilder('PHPExcel_Style_NumberFormat')->disableOriginalConstructor()->getMock();
    }
}