<?php

namespace ExcelAnt\PhpExcel\Writer;

use PHPExcel_Worksheet;

use ExcelAnt\Table\Table;
use ExcelAnt\Cell\EmptyCell;
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
        $tableValues = $table->getTable();

        // Labels here

        // Table handling
        foreach ($tableValues as $row) {
            foreach ($row as $index => $cell) {

                if ($cell instanceof EmptyCell) {
                    $coordinate->nextXAxis();

                    continue;
                }

                $cellFormat = null;
                $styleCollection = $cell->getStyles();

                if (!empty($styleCollection)) {
                    try {
                        $cellFormat = $styleCollection->getElement(new Format())->getFormat();
                    } catch (\OutOfBoundsException $e) {}
                }

                $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), $cell->getValue(), $cellFormat);

                if ($cell->hasStyles()) {
                    $phpExcelWorksheet = $this->styleWorker->applyStyles($phpExcelWorksheet, $coordinate, $styleCollection);
                }

                $coordinate->nextXAxis();
            }

            $coordinate->resetXAxis()->nextYAxis();
        }

        return $phpExcelWorksheet;
    }
}