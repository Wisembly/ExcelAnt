<?php

namespace ExcelAnt\PhpExcel\Writer;

use PHPExcel_Worksheet;

use ExcelAnt\Table\Table;
use ExcelAnt\Cell\EmptyCell;

class TableWorker
{
    public function writeTable(PHPExcel_Worksheet $phpExcelWorksheet, Table $table)
    {
        $coordinate = $table->getCoordinate();
        $tableValues = $table->getTable();

        // Ecrire les labels

        foreach ($tableValues as $row) {
            foreach ($row as $index => $cell) {

                if ($cell instanceof EmptyCell) {
                    $coordinate->nextXAxis();

                    continue;
                }

                $phpExcelWorksheet->setCellValueByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), $cell->getValue());

                // Style here

                $coordinate->nextXAxis();
            }

            $coordinate->resetXAxis();
            $coordinate->nextYAxis();
        }

        return $phpExcelWorksheet;
    }
}