<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\Writer\WriterInterface;
use ExcelAnt\PhpExcel\Writer\TableWorker;
use ExcelAnt\PhpExcel\Writer\CellWorker;
use ExcelAnt\Workbook\WorkbookInterface;

class Writer implements WriterInterface
{
    private $workbook;
    private $tableWorker;
    private $cellWorker;

    /**
     * @param WorkbookInterface $workbook The workbook to be exported
     */
    public function __construct(WorkbookInterface $workbook, TableWorker $tableWorker, CellWorker $cellWorker)
    {
        $this->setWorkbook($workbook);
        $this->tableWorker = $tableWorker;
        $this->cellWorker = $cellWorker;
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
                $this->tableWorker->writeTable($phpExcelWorksheet, $table);
            }

            // Write the individuals cells
            foreach ($sheet->getCells() as $cell) {
                $this->cellWorker->writeCell($cell, $phpExcelWorksheet, $cell->getCoordinate());
            }

            $phpExcel->addSheet($phpExcelWorksheet);
        }
    }
}