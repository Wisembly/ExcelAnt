<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\Writer\WriterInterface;
use ExcelAnt\Workbook\WorkbookInterface;

class Writer implements WriterInterface
{
    private $workbook;

    public function __construct(WorkbookInterface $workbook = null)
    {
        if (isset($workbook)) {
            $this->setWorkbook($workbook);
        }
    }

    public function setWorkbook(WorkbookInterface $workbook)
    {
        $this->workbook = $workbook;

        return $this;
    }

    public function getWorkbook()
    {
        return $this->workbook;
    }

    public function write($path)
    {

    }
}