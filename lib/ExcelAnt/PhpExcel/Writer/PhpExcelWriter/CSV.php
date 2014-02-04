<?php

namespace ExcelAnt\PhpExcel\Writer\PhpExcelWriter;

use PHPExcel;
use PHPExcel_Writer_CSV;

use ExcelAnt\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface;

class CSV implements PhpExcelWriterInterface
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
        $writer = (new PHPExcel_Writer_CSV($phpExcel))
            ->save($this->path);
    }
}