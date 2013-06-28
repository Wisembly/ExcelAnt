<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer;

use ExcelAnt\Adapter\PhpExcel\Writer\Writer;

class WriterFactoryTest extends \PHPUnit_Framework_TestCase
{
    public function testCreateWriter()
    {
        $this->assertInstanceOf('ExcelAnt\Adapter\PhpExcel\Writer\Writer', (new WriterFactory())->createWriter($this->getPhpExcelWriterInterfaceMock()));
    }

    /**
     * Mock PhpExcelWriterInterface
     *
     * @return Mock_PhpExcelWriterInterface
     */
    public function getPhpExcelWriterInterfaceMock()
    {
        return $this->getMockBuilder('ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface')->disableOriginalConstructor()->getMock();
    }
}