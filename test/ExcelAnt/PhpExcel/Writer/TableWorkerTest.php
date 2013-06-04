<?php

namespace ExcelAnt\PhpExcel\Writer;

use PHPExcel_Worksheet;
use PHPExcel_Style;

use ExcelAnt\PhpExcel\Writer\TableWorker;
use ExcelAnt\Table\Table;
use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Collections\StyleCollection;
use ExcelAnt\Style\Fill;
use ExcelAnt\Style\Font;
use ExcelAnt\Style\Format;

class TableWorkerTest extends \PHPUnit_Framework_TestCase
{
    public function testWriteTable()
    {
        $localCellStorage;

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->any())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($pColumn, $pRow, $pValue) use (&$localCellStorage) {
                $localCellStorage[$pRow][$pColumn] = $pValue;
            }));

        $tableWorker = new TableWorker($this->getStyleWorkerMock());
        $response = $tableWorker->writeTable($phpExcelWorksheet, $this->getTable());

        $expectedArray = [
            1 => [
                0 => 'foo',
                3 => 'bar',
                4 => 'baz',
            ],
            2 => [
                0 => 'foo',
                1 => 'bar',
            ],
        ];

        $this->assertEquals($expectedArray, $localCellStorage);
    }

    public function testWriteTableWithASpecificFormat()
    {
        $localCellStorage;

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->any())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnCallback(function($pColumn, $pRow, $pValue, $pDataType) use (&$localCellStorage) {
                $localCellStorage[$pRow][$pColumn] = $pDataType;
            }));

        $styleWorker = $this->getStyleWorkerMock();
        $styleWorker->expects($this->any())
            ->method('applyStyles')
            ->will($this->returnValue($phpExcelWorksheet));

        $tableWorker = new TableWorker($styleWorker);
        $response = $tableWorker->writeTable($phpExcelWorksheet, $this->getTableWithSpecificFormat());

        $expectedArray = [
            1 => [
                0 => null,
                3 => null,
                4 => null,
            ],
            2 => [
                0 => Format::TYPE_NUMERIC,
                1 => Format::TYPE_NUMERIC,
            ],
        ];

        $this->assertEquals($expectedArray, $localCellStorage);
    }

    public function testWriteTableWithStyles()
    {
        $localApplyStyles = [];

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();

        $phpExcelWorksheet->expects($this->any())
            ->method('setCellValueExplicitByColumnAndRow')
            ->will($this->returnValue(true));

        $styleWorker = $this->getStyleWorkerMock();

        $styleWorker->expects($this->atLeastOnce())
            ->method('applyStyles')
            ->will($this->returnCallback(function($worksheet, $coordinate, $styleCollection) use (&$localApplyStyles, &$phpExcelWorksheet) {
                $localApplyStyles[] = [
                    'xAxis' => $coordinate->getXAxis(),
                    'yAxis' => $coordinate->getYAxis(),
                ];

                return $phpExcelWorksheet;
            }));

        $expected = [
            ['xAxis' => 1, 'yAxis' => 2],
            ['xAxis' => 2, 'yAxis' => 2],
        ];

        $tableWorker = new TableWorker($styleWorker);
        $response = $tableWorker->writeTable($phpExcelWorksheet, $this->getTableWithStyles());

        $this->assertEquals($expected, $localApplyStyles);
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
     * Mock StyleWorker
     * @return Mock
     */
    private function getStyleWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\PhpExcel\Writer\StyleWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Get a Table with data
     *
     * @return Table
     */
    private function getTable()
    {
        $table = (new Table())
            ->setRow(['foo', null, null, 'bar', 'baz', null])
            ->setRow(['foo', 'bar'])
            ->setCoordinate(new Coordinate(1, 1));

        return $table;
    }

    /**
     * Get a Table with styles
     *
     * @return Table
     */
    private function getTableWithStyles()
    {
        $table = (new Table())
            ->setRow(['foo', null, null, 'bar', 'baz', null])
            ->setRow(['foo', 'bar'], null, new StyleCollection([new Fill(), new Font()]))
            ->setCoordinate(new Coordinate(1, 1));

        return $table;
    }

    /**
     * Get a Table with specific format
     *
     * @return Table
     */
    private function getTableWithSpecificFormat()
    {
        $table = (new Table())
            ->setRow(['foo', null, null, 'bar', 'baz', null])
            ->setRow(['foo', 'bar'], null, new StyleCollection([(new Format())->setFormat(Format::TYPE_NUMERIC)]))
            ->setCoordinate(new Coordinate(1, 1));

        return $table;
    }
}