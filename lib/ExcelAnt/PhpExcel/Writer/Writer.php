<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\Writer\WriterInterface;
use ExcelAnt\PhpExcel\Writer\Worker\TableWorker;
use ExcelAnt\PhpExcel\Writer\Worker\CellWorker;
use ExcelAnt\Workbook\WorkbookInterface;

class Writer implements WriterInterface
{
    private $tableWorker;
    private $cellWorker;

    /**
     * @param WorkbookInterface $workbook The workbook to be exported
     */
    public function __construct(TableWorker $tableWorker, CellWorker $cellWorker)
    {
        $this->tableWorker = $tableWorker;
        $this->cellWorker = $cellWorker;
    }

    public function write(WorkbookInterface $workbook, $path)
    {
        $phpExcel = $workbook->getRawClass();

        foreach ($workbook->getAllSheets() as $sheet) {
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