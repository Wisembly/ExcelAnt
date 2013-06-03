<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\Writer\WriterInterface;
use ExcelAnt\PhpExcel\Writer\TableWorker;
use ExcelAnt\Workbook\WorkbookInterface;

class Writer implements WriterInterface
{
    private $workbook;
    private $tableWorker;

    public function __construct(WorkbookInterface $workbook = null)
    {
        if (isset($workbook)) {
            $this->setWorkbook($workbook);
        }

        $this->tableWorker = new TableWorker();
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
        $phpExcel = $this->workbook->getRawClass();

        foreach ($this->workbook->getAllSheets() as $sheet) {
            $phpExcelWorksheet = $sheet->getRawClass();

            foreach ($sheet->getTables() as $table) {
                $phpExcelWorksheet = $this->tableWorker->writeTable($phpExcelWorksheet, $table);
            }

            $phpExcel->addSheet($phpExcelWorksheet);
        }
    }
}