<?php

namespace ExcelAnt\Writer;

use ExcelAnt\Workbook\WorkbookInterface;

interface WriterInterface
{
    /**
     * Set a Workbook
     *
     * @param WorkbookInterface $workbook
     *
     * @return WriterInterface
     */
    public function setWorkbook(WorkbookInterface $workbook);

    /**
     * get the associated Workbook
     *
     * @return WorkbookInterface
     */
    public function getWorkbook();

    public function write($path);
}