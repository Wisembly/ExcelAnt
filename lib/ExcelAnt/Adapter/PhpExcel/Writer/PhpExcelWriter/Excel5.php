<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter;

use \PHPExcel;
use \PHPExcel_Writer_Excel5;

use ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface;

class Excel5 implements PhpExcelWriterInterface
{
    private $path;

    /**
     * @param string $path The location where the file will be saved
     */
    public function __construct($path)
    {
        $this->path = $path;
    }

    /**
     * {@inheritdoc}
     */
    public function save(PHPExcel $phpExcel)
    {
        $writer = (new PHPExcel_Writer_Excel5($phpExcel))
            ->save($this->path);
    }
}