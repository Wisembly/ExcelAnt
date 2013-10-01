<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter;

use \PHPExcel;
use \PHPExcel_Writer_Excel2007;

class Excel2007 implements PhpExcelWriterInterface
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
        $writer = (new PHPExcel_Writer_Excel2007($phpExcel))
            ->save($this->path);
    }
}