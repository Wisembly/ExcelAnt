<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\Worker;

use ExcelAnt\Adapter\PhpExcel\Writer\Worker\LabelWorker,
    ExcelAnt\Table\Label,
    ExcelAnt\Coordinate\Coordinate;

class LabelWorkerTest extends \PHPUnit_Framework_TestCase
{
    public function testWriteTopLabel()
    {
        $localWriteCellStorage = [];

        $label = new Label(Label::TOP);
        $label->setValues(['foo', 'bar', 'baz']);

        $cellWorker = $this->getCellWorkerMock();
        $cellWorker->expects($this->exactly(3))
            ->method('writeCell')
            ->will($this->returnCallback(function($cell, $phpExcelWorsheet, $coordinate) use (&$localWriteCellStorage) {
                $localWriteCellStorage[] = [$cell->getValue(), $coordinate->getXAxis()];
            }));

        $coordinate = new Coordinate(1, 1);

        $labelWorker = (new LabelWorker($cellWorker))
            ->writeLabel($label, $this->getPhpExcelWorksheetMock(), $coordinate);

        $this->assertEquals(1, $coordinate->getXAxis());
        $this->assertEquals(2, $coordinate->getYAxis());
        $this->assertEquals(1, $coordinate->getOriginalXAxis());
        $this->assertEquals(2, $coordinate->getOriginalYAxis());

        $expected = [
            ['foo', 1],
            ['bar', 2],
            ['baz', 3],
        ];

        $this->assertEquals($expected, $localWriteCellStorage);
    }

    public function testWriteLeftLabel()
    {
        $localWriteCellStorage = [];

        $label = new Label(Label::LEFT);
        $label->setValues(['foo', 'bar', 'baz']);

        $cellWorker = $this->getCellWorkerMock();
        $cellWorker->expects($this->exactly(3))
            ->method('writeCell')
            ->will($this->returnCallback(function($cell, $phpExcelWorsheet, $coordinate) use (&$localWriteCellStorage) {
                $localWriteCellStorage[] = [$cell->getValue(), $coordinate->getYAxis()];
            }));

        $coordinate = new Coordinate(1, 1);

        $labelWorker = (new LabelWorker($cellWorker))
            ->writeLabel($label, $this->getPhpExcelWorksheetMock(), $coordinate);

        $this->assertEquals(2, $coordinate->getXAxis());
        $this->assertEquals(1, $coordinate->getYAxis());
        $this->assertEquals(2, $coordinate->getOriginalXAxis());
        $this->assertEquals(1, $coordinate->getOriginalYAxis());

        $expected = [
            ['foo', 1],
            ['bar', 2],
            ['baz', 3],
        ];

        $this->assertEquals($expected, $localWriteCellStorage);
    }

    public function testWriteFullLabel()
    {
        $localWriteCellStorage = [];

        $label = new Label(Label::FULL);
        $label->setValues([['foo', 'bar', 'baz'], ['foofoo', 'barbar', 'bazbaz']]);

        $cellWorker = $this->getCellWorkerMock();
        $cellWorker->expects($this->exactly(6))
            ->method('writeCell')
            ->will($this->returnCallback(function($cell, $phpExcelWorsheet, $coordinate) use (&$localWriteCellStorage) {
                $localWriteCellStorage[] = [$cell->getValue(), $coordinate->getXAxis(), $coordinate->getYAxis()];
            }));

        $coordinate = new Coordinate(1, 1);

        $labelWorker = (new LabelWorker($cellWorker))
            ->writeLabel($label, $this->getPhpExcelWorksheetMock(), $coordinate);

        $this->assertEquals(2, $coordinate->getXAxis());
        $this->assertEquals(2, $coordinate->getYAxis());
        $this->assertEquals(2, $coordinate->getOriginalXAxis());
        $this->assertEquals(2, $coordinate->getOriginalYAxis());

        $expected = [
            ['foo', 2, 1],
            ['bar', 3, 1],
            ['baz', 4, 1],
            ['foofoo', 1, 2],
            ['barbar', 1, 3],
            ['bazbaz', 1, 4],
        ];

        $this->assertEquals($expected, $localWriteCellStorage);
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
}