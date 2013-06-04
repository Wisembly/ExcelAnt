<?php

namespace ExcelAnt\PhpExcel\Writer;

use PHPExcel_Worksheet;

use ExcelAnt\Table\Table;
use ExcelAnt\Table\Label;
use ExcelAnt\Cell\EmptyCell;
use ExcelAnt\Cell\CellInterface;
use ExcelAnt\PhpExcel\Writer\StyleWorker;
use ExcelAnt\Style\Format;

class TableWorker
{
    private $styleWorker;

    public function __construct(StyleWorker $styleWorker)
    {
        $this->styleWorker = $styleWorker;
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
            if (Label::TOP === $label->getType()) {
                foreach ($label->getValues() as $cell) {

                    if ($cell->hasStyles()) {
                        $phpExcelWorksheet = $this->styleWorker->applyStyles($phpExcelWorksheet, $coordinate, $cell->getStyles());
                    }

                    if ($cell instanceof EmptyCell) {
                        $coordinate->nextXAxis();

                        continue;
                    }

                    $cellFormat = $this->getCellFormat($cell);
                    $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), $cell->getValue(), $cellFormat);
                    $coordinate->nextXAxis();
                }

                $coordinate->resetXAxis()->nextYAxis();
            }

            // TODO les autres types de labels
        }

        // Table handling
        foreach ($table->getTable() as $row) {
            foreach ($row as $index => $cell) {

                if ($cell->hasStyles()) {
                    $phpExcelWorksheet = $this->styleWorker->applyStyles($phpExcelWorksheet, $coordinate, $cell->getStyles());
                }

                if ($cell instanceof EmptyCell) {
                    $coordinate->nextXAxis();

                    continue;
                }

                $cellFormat = $this->getCellFormat($cell);
                $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), $cell->getValue(), $cellFormat);
                $coordinate->nextXAxis();
            }

            $coordinate->resetXAxis()->nextYAxis();
        }

        return $phpExcelWorksheet;
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