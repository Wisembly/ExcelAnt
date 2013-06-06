<?php

namespace ExcelAnt\PhpExcel\Writer\Worker;

use ExcelAnt\PhpExcel\Writer\Worker\TableWorker;
use ExcelAnt\PhpExcel\Writer\Worker\LabelWorker;
use ExcelAnt\PhpExcel\Writer\Worker\StyleWorker;
use ExcelAnt\PhpExcel\Writer\Worker\CellWorker;

class WorkerFactory
{
    public function createTableWorker()
    {
        return new TableWorker($this->createCellWorker(), $this->createLabelWorker());
    }

    public function createLabelWorker()
    {
        return new LabelWorker($this->createCellWorker());
    }

    public function createStyleWorker()
    {
        return new StyleWorker();
    }

    public function createCellWorker()
    {
        return new CellWorker($this->createStyleWorker());
    }
}