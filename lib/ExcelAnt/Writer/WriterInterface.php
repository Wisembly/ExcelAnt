<?php

namespace ExcelAnt\Writer;

use PHPExcel;

use ExcelAnt\Workbook\WorkbookInterface;
use ExcelAnt\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface;

interface WriterInterface
{
    /**
     * Convert your Workbook into PHPExcel class
     *
     * @param WorkbookInterface $workbook
     *
     * @return PHPExcel
     */
    public function convert(WorkbookInterface $workbook);

    /**
     * Write your Workbook
     *
     * @param PHPExcel $phpExcel
     */
    public function write(PHPExcel $phpExcel);
}