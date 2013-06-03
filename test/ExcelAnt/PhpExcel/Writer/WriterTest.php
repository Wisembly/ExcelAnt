<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\PhpExcel\Writer\Writer;
use ExcelAnt\PhpExcel\Workbook;
use ExcelAnt\PhpExcel\Sheet;
use ExcelAnt\Table\Table;
use ExcelAnt\Coordinate\Coordinate;

class WriterTest extends \PHPUnit_Framework_TestCase
{
    public function testInstanciateWriter()
    {
        $workbook = new Workbook();
        $writer = new Writer($workbook);

        $this->assertInstanceOf('ExcelAnt\PhpExcel\Workbook', $writer->getWorkbook());
    }

    public function testSetAndGetWriter()
    {
        $workbook = (new Workbook())->setTitle('Foo');
        $writer = new Writer($workbook);
        $writer->setWorkbook($workbook);

        $this->assertInstanceOf('ExcelAnt\PhpExcel\Workbook', $writer->getWorkbook());
        $this->assertEquals('Foo', $writer->getWorkbook()->getTitle());
    }
}