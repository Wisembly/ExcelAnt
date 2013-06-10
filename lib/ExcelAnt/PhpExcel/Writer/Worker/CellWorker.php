<?php

namespace ExcelAnt\PhpExcel\Writer\Worker;

use \PHPExcel_Worksheet;

use ExcelAnt\PhpExcel\Writer\Worker\StyleWorker;
use ExcelAnt\Cell\CellInterface;
use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Cell\EmptyCell;
use ExcelAnt\Style\Format;

class CellWorker
{
    private $styleWorker;

    public function __construct(StyleWorker $styleWorker)
    {
        $this->styleWorker = $styleWorker;
    }

    public function writeCell(CellInterface $cell, PHPExcel_Worksheet $phpExcelWorksheet, Coordinate $coordinate)
    {
        if ($cell->hasStyles()) {
            $styles = $this->styleWorker->convertStyles($cell->getStyles());

            if (!empty($styles)) {
                $phpExcelWorksheet
                    ->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())
                    ->applyFromArray($styles);
            }
        }

        if ($cell instanceof EmptyCell) {
            return;
        }

        $cellFormat = $this->getCellFormat($cell);
        $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), $cell->getValue(), $cellFormat);
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

        return Format::TYPE_STRING;
    }
}