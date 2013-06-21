<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\Worker;

use \PHPExcel_Worksheet,
    \PHPExcel_Style_NumberFormat;

use ExcelAnt\Adapter\PhpExcel\Writer\Worker\StyleWorker,
    ExcelAnt\Cell\CellInterface,
    ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Cell\EmptyCell,
    ExcelAnt\Style\Format;

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

        $styleCollection = $cell->getStyles();

        if (!empty($styleCollection)) {
            try {
                $cellFormat = $styleCollection->getElement(new Format())->getFormat();

                if (Format::TYPE_NUMERIC === $cellFormat) {
                    $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
                } elseif (Format::TYPE_PERCENT === $cellFormat) {
                    $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);
                } elseif (Format::TYPE_DATETIME === $cellFormat) {
                    $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
                }

            } catch (\OutOfBoundsException $e) {}
        }

        $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), $cell->getValue());
    }
}