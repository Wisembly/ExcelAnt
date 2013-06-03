<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\PhpExcel\Writer\Writer;
use ExcelAnt\PhpExcel\Workbook;

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