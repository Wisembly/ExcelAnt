<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\Writer\WriterInterface;
use ExcelAnt\PhpExcel\Writer\TableWorker;
use ExcelAnt\Workbook\WorkbookInterface;

class Writer implements WriterInterface
{
    private $workbook;
    private $tableWorker;

    /**
     * @param WorkbookInterface $workbook The workbook to be exported
     */
    public function __construct(WorkbookInterface $workbook)
    {
        $this->setWorkbook($workbook);
        $this->tableWorker = new TableWorker();
    }

    /**
     * Replace workbook
     *
     * @param WorkbookInterface $workbook
     *
     * @return Writer
     */
    public function setWorkbook(WorkbookInterface $workbook)
    {
        $this->workbook = $workbook;

        return $this;
    }

    /**
     * Get associated workbook
     *
     * @return WorkbookInterface
     */
    public function getWorkbook()
    {
        return $this->workbook;
    }

    public function write($path)
    {
        $phpExcel = $this->workbook->getRawClass();

        foreach ($this->workbook->getAllSheets() as $sheet) {
            $phpExcelWorksheet = $sheet->getRawClass();

            // Write the tables
            foreach ($sheet->getTables() as $table) {
                $phpExcelWorksheet = $this->tableWorker->writeTable($phpExcelWorksheet, $table);
            }

            // Write the individual cells
            // TODO

            $phpExcel->addSheet($phpExcelWorksheet);
        }
    }
}