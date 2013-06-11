<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer\Worker;

use \PHPExcel_Worksheet;

use ExcelAnt\Table\LabelInterface,
    ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\CellWorker;

class LabelWorker
{
    private $cellWorker;

    public function __construct(CellWorker $cellWorker)
    {
        $this->cellWorker = $cellWorker;
    }

    public function writeLabel(LabelInterface $label, PHPExcel_Worksheet $phpExcelWorksheet, Coordinate $coordinate)
    {
        $values = $label->getValues();

        switch ($label->getType()) {
            case LabelInterface::TOP:
                    $this->writeTopLabel($values, $phpExcelWorksheet, $coordinate);
                break;

            case LabelInterface::LEFT:
                    $this->writeLeftLabel($values, $phpExcelWorksheet, $coordinate);
                break;

            case LabelInterface::FULL:
                    $this->writeFullLabel($values, $phpExcelWorksheet, $coordinate);
                break;
        }
    }

    private function writeTopLabel(array $values, PHPExcel_Worksheet $phpExcelWorksheet, Coordinate $coordinate)
    {
        foreach ($values as $cell) {
            $this->cellWorker->writeCell($cell, $phpExcelWorksheet, $coordinate);
            $coordinate->nextXAxis();
        }

        $coordinate->setOriginalYAxis($coordinate->getOriginalYAxis() + 1)
            ->resetYAxis()
            ->resetXAxis();
    }

    private function writeLeftLabel(array $values, PHPExcel_Worksheet $phpExcelWorksheet, Coordinate $coordinate)
    {
        foreach ($values as $cell) {
            $this->cellWorker->writeCell($cell, $phpExcelWorksheet, $coordinate);
            $coordinate->nextYAxis();
        }

        $coordinate->setOriginalXAxis($coordinate->getOriginalXAxis() + 1)
            ->resetYAxis()
            ->resetXAxis();
    }

    private function writeFullLabel(array $values, PHPExcel_Worksheet $phpExcelWorksheet, Coordinate $coordinate)
    {
        for ($i = 0; $i < 2; $i++) {
            if (0 === $i) {
                $coordinate->setXAxis($coordinate->getOriginalXAxis() + 1);
            } elseif (1 === $i) {
                $coordinate->resetXAxis();
                $coordinate->setYAxis($coordinate->getOriginalYAxis() + 1);
            }

            $lengthSide = count($values[$i]);

            for ($j = 0; $j < $lengthSide; $j++) {
                $this->cellWorker->writeCell($values[$i][$j], $phpExcelWorksheet, $coordinate);

                if (0 === $i) {
                    $coordinate->nextXAxis();
                } elseif (1 === $i) {
                    $coordinate->nextYAxis();
                }
            }
        }

        $coordinate->setOriginalXAxis($coordinate->getOriginalXAxis() + 1)
            ->setOriginalYAxis($coordinate->getOriginalYAxis() + 1)
            ->resetYAxis()
            ->resetXAxis();
    }
}