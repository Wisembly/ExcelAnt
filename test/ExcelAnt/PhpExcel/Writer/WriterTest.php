<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\PhpExcel\Writer\Writer;
use ExcelAnt\PhpExcel\Workbook;
use ExcelAnt\PhpExcel\Sheet;
use ExcelAnt\Table\Table;
use ExcelAnt\Coordinate\Coordinate;

class WriterTest extends \PHPUnit_Framework_TestCase
{
    public function testSetAndGetWriter()
    {
        $workbook = new Workbook();
        $writer = new Writer();
        $writer->setWorkbook($workbook);

        $this->assertInstanceOf('ExcelAnt\PhpExcel\Workbook', $writer->getWorkbook());
    }
}