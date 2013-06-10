<?php

namespace ExcelAnt\Adapter\PhpExcel\Workbook;

use ExcelAnt\Workbook\WorkbookInterface;

interface WorkbookExcelInterface extends WorkbookInterface
{
    /**
     * Get raw class
     *
     * @return mixed
     */
    public function getRawClass();
}