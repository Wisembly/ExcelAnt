<?php

namespace ExcelAnt\Writer;

use ExcelAnt\Workbook\WorkbookInterface;
use ExcelAnt\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface;

interface WriterInterface
{
    /**
     * Write your Workbook
     *
     * @param  WorkbookInterface       $workbook
     */
    public function write(WorkbookInterface $workbook);
}