<?php

namespace ExcelAnt\PhpExcel\Writer;

use PHPExcel_Worksheet;

use ExcelAnt\Table\Table;
use ExcelAnt\Table\Label;
use ExcelAnt\Cell\EmptyCell;
use ExcelAnt\Cell\CellInterface;
use ExcelAnt\PhpExcel\Writer\CellWorker;
use ExcelAnt\PhpExcel\Writer\LabelWorker;
use ExcelAnt\Style\Format;

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
            $this->labelWorker->writeLabel($phpExcelWorksheet, $label, $coordinate);
        }

        // Table handling
        foreach ($table->getTable() as $row) {
            foreach ($row as $index => $cell) {
                $this->cellWorker->writeCell($phpExcelWorksheet, $label, $cell);
            }

            $coordinate->resetXAxis()->nextYAxis();
        }
    }

    /**
     * Get the cell format.
     *
     * @param  CellInterface $cell
     *
     * @return mixed The format as string or null
     */
    private function getCellFormat(CellInterface $cell)
    {
        $styleCollection = $cell->getStyles();

        if (!empty($styleCollection)) {
            try {
                return $styleCollection->getElement(new Format())->getFormat();
            } catch (\OutOfBoundsException $e) {}
        }

        return null;
    }
}