<?php

namespace ExcelAnt\Adapter\PhpExcel\Writer;

use ExcelAnt\Adapter\PhpExcel\Writer\Worker\StyleWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\CellWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\LabelWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\TableWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\PhpExcelWriterInterface,
    ExcelAnt\Adapter\PhpExcel\Writer\Writer;

class WriterFactory
{
    public function createWriter(PhpExcelWriterInterface $writer)
    {
        $styleWorker = new StyleWorker();
        $cellWorker = new CellWorker($styleWorker);
        $labelWorker = new LabelWorker($cellWorker);
        $tableWorker = new TableWorker($cellWorker, $labelWorker);

        return new Writer($writer, $tableWorker, $cellWorker, $styleWorker);
    }
}