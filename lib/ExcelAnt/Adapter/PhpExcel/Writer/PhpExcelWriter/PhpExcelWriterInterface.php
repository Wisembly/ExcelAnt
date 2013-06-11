<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter;

use \PHPExcel;

interface PhpExcelWriterInterface
{
    /**
     * @param  PHPExcel $phpExcel
     */
    public function save(PHPExcel $phpExcel);
}