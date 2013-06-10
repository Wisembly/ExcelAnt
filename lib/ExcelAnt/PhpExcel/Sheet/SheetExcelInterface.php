<?php

namespace ExcelAnt\PhpExcel\Sheet;

use ExcelAnt\Sheet\SheetInterface;

interface SheetExcelInterface extends SheetInterface
{
    /**
     * Get raw class
     *
     * @return mixed
     */
    public function getRawClass();
}