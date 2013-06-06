<?php

namespace ExcelAnt\PhpExcel\Writer;

use ExcelAnt\Writer\WriterInterface;
use ExcelAnt\PhpExcel\Writer\Worker\TableWorker;
use ExcelAnt\PhpExcel\Writer\Worker\CellWorker;
use ExcelAnt\PhpExcel\Writer\Worker\StyleWorker;
use ExcelAnt\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface;
use ExcelAnt\Workbook\WorkbookInterface;

class Writer implements WriterInterface
{
    private $tableWorker;
    private $cellWorker;
    private $styleWorker;

    /**
     * @param WorkbookInterface $workbook The workbook to be exported
     */
    public function __construct(TableWorker $tableWorker, CellWorker $cellWorker, StyleWorker $styleWorker)
    {
        $this->tableWorker = $tableWorker;
        $this->cellWorker = $cellWorker;
        $this->styleWorker = $styleWorker;
    }

    /**
     * {@inheritdoc}
     */
    public function write(WorkbookInterface $workbook, PhpExcelWriterInterface $writer)
    {
        $phpExcel = $workbook->getRawClass();

        // Apply default workbook styles
        if ($workbook->hasStyles()) {
            $styles = $this->styleWorker->convertStyles($workbook->getStyles());

            $phpExcel->getDefaultStyle()->applyFromArray($styles);
        }

        foreach ($workbook->getAllSheets() as $sheet) {
            $phpExcelWorksheet = $sheet->getRawClass();

            $phpExcel->addSheet($phpExcelWorksheet);

            // Write the tables
            foreach ($sheet->getTables() as $table) {
                $this->tableWorker->writeTable($phpExcelWorksheet, $table);
            }

            // Write the individuals cells
            foreach ($sheet->getCells() as $cell) {
                $this->cellWorker->writeCell($cell, $phpExcelWorksheet, $cell->getCoordinate());
            }
        }

        $writer->save($phpExcel);
    }
}