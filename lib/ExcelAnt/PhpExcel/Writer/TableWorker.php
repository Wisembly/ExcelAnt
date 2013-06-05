<?php

namespace ExcelAnt\PhpExcel\Writer;

use PHPExcel_Worksheet;

use ExcelAnt\Table\Table;
use ExcelAnt\PhpExcel\Writer\CellWorker;
use ExcelAnt\PhpExcel\Writer\LabelWorker;

class TableWorker
{
    private $cellWorker;
    private $labelWorker;

    public function __construct(CellWorker $cellWorker, LabelWorker $labelWorker)
    {
        $this->cellWorker = $cellWorker;
        $this->labelWorker = $labelWorker;
    }

    /**
     * Convert Table to PHPExcel_Worksheet data
     *
     * @param  PHPExcel_Worksheet $phpExcelWorksheet The current worksheet of the workbook
     * @param  Table              $table             The Table to convert
     *
     * @return PHPExcel_Worksheet
     */
    public function writeTable(PHPExcel_Worksheet $phpExcelWorksheet, Table $table)
    {
        $coordinate = $table->getCoordinate();

        // Labels handling
        if (null !== $label = $table->getLabel()) {
            $this->labelWorker->writeLabel($label, $phpExcelWorksheet, $coordinate);
        }

        // Table handling
        foreach ($table->getTable() as $row) {
            foreach ($row as $index => $cell) {
                $this->cellWorker->writeCell($cell, $phpExcelWorksheet, $coordinate);
                $coordinate->nextXAxis();
            }

            $coordinate->resetXAxis()->nextYAxis();
        }
    }
}