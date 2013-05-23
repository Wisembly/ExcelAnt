<?php

namespace ExcelAnt\PhpExcel;

use PHPExcel_Worksheet;

use ExcelAnt\Sheet\SheetInterface;

class Sheet implements SheetInterface
{
    private $phpExcelWorksheet;

    public function __construct(Worksheet $worksheet)
    {
        $this->phpExcelWorksheet = new PHPExcel_Worksheet($worksheet->getRawClass());
    }

    public function getRawClass()
    {
        return $this->phpExcelWorksheet;
    }

    public function setTitle($name)
    {
        $this->phpExcelWorksheet->setTitle($name);

        return $this;
    }

    public function getTitle()
    {
        return $this->phpExcelWorksheet->getTitle();
    }

    public function writeCell()
    {

    }

    public function addTable()
    {

    }

    public function setRowHieght()
    {

    }

    public function setColumnWidth()
    {

    }

    public function importImage()
    {

    }

    public function write()
    {

    }
}