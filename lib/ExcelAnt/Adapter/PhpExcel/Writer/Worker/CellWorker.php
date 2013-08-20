<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\Worker;

use \PHPExcel_Worksheet,
    \PHPExcel_Style_NumberFormat,
    \PHPExcel_Shared_Date,
    \PHPExcel_Cell,
    \PHPExcel_Cell_AdvancedValueBinder,
    \PHPExcel_Cell_DataType;

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
        $value = $cell->getValue();

        if (!empty($styleCollection)) {
            try {

                switch ($styleCollection->getElement(new Format())->getFormat()) {
                    case Format::TYPE_NUMERIC:
                        $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
                        $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), (int) $value, PHPExcel_Cell_DataType::TYPE_NUMERIC);

                        return;
                        break;

                    case Format::TYPE_NUMERIC_00:
                        $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
                        $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), (float) $value, PHPExcel_Cell_DataType::TYPE_NUMERIC);

                        return;
                        break;

                    case Format::TYPE_PERCENT_00:
                        $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
                        $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), (float) $value, PHPExcel_Cell_DataType::TYPE_NUMERIC);

                        return;
                        break;

                    case Format::TYPE_PERCENT:
                        $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);
                        $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), (float) $value, PHPExcel_Cell_DataType::TYPE_NUMERIC);

                        return;
                        break;

                    case Format::TYPE_DATE:
                        $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DDMMYYYY);

                        break;

                    case Format::TYPE_DATETIME:
                        $phpExcelWorksheet->getStyleByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis())->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);

                        break;
                }
            } catch (\OutOfBoundsException $e) {}
        }

        $phpExcelWorksheet->setCellValueExplicitByColumnAndRow($coordinate->getXAxis() - 1, $coordinate->getYAxis(), $value);
    }
}