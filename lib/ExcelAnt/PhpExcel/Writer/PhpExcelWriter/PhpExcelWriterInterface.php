<?php

namespace ExcelAnt\PhpExcel\Writer\PhpExcelWriter;

use \PHPExcel;

interface PhpExcelWriterInterface
{
    /**
     * @param  PHPExcel $phpExcel
     */
    public function save(PHPExcel $phpExcel);
}