<?php

namespace ExcelAnt\PhpExcel\Writer;

use PHPExcel_Worksheet;

use ExcelAnt\Table\Table;
use ExcelAnt\Cell\EmptyCell;

class TableWorker
{
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

                $phpExcelWorksheet->setCellValueByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), $cell->getValue());

                if ($cell->hasStyles()) {
                    // Styles here
                }

                $coordinate->nextXAxis();
            }

            $coordinate->resetXAxis()->nextYAxis();
        }

        return $phpExcelWorksheet;
    }
}