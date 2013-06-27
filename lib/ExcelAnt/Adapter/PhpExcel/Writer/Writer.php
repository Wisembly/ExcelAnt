<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer;

use ExcelAnt\Writer\WriterInterface,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\TableWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\CellWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\StyleWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface,
    ExcelAnt\Workbook\WorkbookInterface;

class Writer implements WriterInterface
{
    private $tableWorker;
    private $cellWorker;
    private $styleWorker;
    private $writer;

    /**
     * @param WorkbookInterface $workbook The workbook to be exported
     */
    public function __construct(PhpExcelWriterInterface $writer, TableWorker $tableWorker, CellWorker $cellWorker, StyleWorker $styleWorker)
    {
        $this->tableWorker = $tableWorker;
        $this->cellWorker = $cellWorker;
        $this->styleWorker = $styleWorker;
        $this->writer = $writer;
    }

    /**
     * {@inheritdoc}
     */
    public function convert(WorkbookInterface $workbook)
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

        $phpExcel->setActiveSheetIndex(0);

        return $phpExcel;
    }

    /**
     * {@inheritdoc}
     */
    public function write($phpExcel)
    {
        $this->writer->save($phpExcel);
    }
}