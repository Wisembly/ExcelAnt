<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\Worker;

use \PHPExcel_Worksheet;
use \PHPExcel_Style;

use ExcelAnt\Adapter\PhpExcel\Writer\Worker\TableWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\LabelWorker,
    ExcelAnt\Table\Table,
    ExcelAnt\Table\Label,
    ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Font,
    ExcelAnt\Style\Format;

class TableWorkerTest extends \PHPUnit_Framework_TestCase
{
    public function testWriteTableWithLabelsAndWithoutData()
    {
        $labelWorker = $this->getLabelWorkerMock();
        $labelWorker->expects($this->once())
            ->method('writeLabel');

        $tableWorker = (new TableWorker($this->getCellWorkerMock(), $labelWorker))
            ->writeTable($this->getPhpExcelWorksheetMock(), $this->getTableWithLabels());
    }

    public function testWriteTableWithLabelsAndData()
    {
        $localCoordinateStorage = [];

        $labelWorker = $this->getLabelWorkerMock();
        $labelWorker->expects($this->once())
            ->method('writeLabel');

        $cellWorker = $this->getCellWorkerMock();
        $cellWorker->expects($this->exactly(2))
            ->method('writeCell');

        $tableWorker = (new TableWorker($cellWorker, $labelWorker))
            ->writeTable($this->getPhpExcelWorksheetMock(), $this->getTableWithLabelsAndData());
    }

    public function testWriteTableWithoutLabels()
    {
        $cellWorker = $this->getCellWorkerMock();
        $cellWorker->expects($this->exactly(8))
            ->method('writeCell');

        $tableWorker = (new TableWorker($cellWorker, $this->getLabelWorkerMock()))
            ->writeTable($this->getPhpExcelWorksheetMock(), $this->getTable());
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
     * Mock CellWorker
     *
     * @return Mock
     */
    private function getCellWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Writer\Worker\CellWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Mock LabelWorker
     *
     * @return Mock
     */
    private function getLabelWorkerMock()
    {
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Writer\Worker\LabelWorker')->disableOriginalConstructor()->getMock();
    }

    /**
     * Get a simple table with data
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
     * Get a table with labels
     *
     * @return Table
     */
    private function getTableWithLabels()
    {
        $label = (new Label())->setValues(['foo', null, 'bar']);

        $table = (new Table())
            ->setLabel($label)
            ->setCoordinate(new Coordinate(1, 1));

        return $table;
    }

    /**
     * Get a table with labels and data
     *
     * @return Table
     */
    private function getTableWithLabelsAndData()
    {
        $label = (new Label())->setValues(['foo', 'bar']);

        $table = (new Table())
            ->setLabel($label)
            ->setRow(['foofoo', 'barbar'])
            ->setCoordinate(new Coordinate(1, 1));

        return $table;
    }
}