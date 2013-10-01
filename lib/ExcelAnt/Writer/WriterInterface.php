<?php

namespace ExcelAnt\Writer;

use ExcelAnt\Workbook\WorkbookInterface;

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
     */
    public function write($classToWrite);
}