<?php

namespace ExcelAnt\PhpExcel\Writer;

use PHPExcel_Worksheet;

use ExcelAnt\PhpExcel\Writer\TableWorker;
use ExcelAnt\Table\Table;
use ExcelAnt\Coordinate\Coordinate;

class TableWorkerTest extends \PHPUnit_Framework_TestCase
{
    private $tableWorker;

    public function setUp()
    {
        $this->tableWorker = new TableWorker();
    }

    public function testWriteTable()
    {
        $localCellStorage;

        $phpExcelWorksheet = $this->getPhpExcelWorksheetMock();
        $phpExcelWorksheet->expects($this->any())
            ->method('setCellValueByColumnAndRow')
            ->will($this->returnCallback(function($pColumn, $pRow, $pValue) use (&$localCellStorage) {
                $localCellStorage[$pRow][$pColumn] = $pValue;
            }));

        $response = $this->tableWorker->writeTable($phpExcelWorksheet, $this->getTable());

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

    /**
     * Mock PHPExcel_Worksheet
     * @return Mock
     */
    private function getPhpExcelWorksheetMock()
    {
        return $this->getMockBuilder('PHPExcel_Worksheet')->disableOriginalConstructor()->getMock();
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
}