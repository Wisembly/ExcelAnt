<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter;

use \PHPExcel,
    \PHPExcel_Writer_HTML;

class Html implements PhpExcelWriterInterface
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
        $writer = (new PHPExcel_Writer_HTML($phpExcel))
            ->save($this->path);
    }
}