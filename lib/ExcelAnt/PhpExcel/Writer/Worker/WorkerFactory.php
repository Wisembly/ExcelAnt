<?php

namespace ExcelAnt\PhpExcel\Writer\Worker;

use ExcelAnt\PhpExcel\Writer\Worker\TableWorker;
use ExcelAnt\PhpExcel\Writer\Worker\LabelWorker;
use ExcelAnt\PhpExcel\Writer\Worker\StyleWorker;
use ExcelAnt\PhpExcel\Writer\Worker\CellWorker;

class WorkerFactory
{
    private $styleWorker;
    private $cellWorker;
    private $labelWorker;
    private $tableWorker;

    public function __construct()
    {
        $this->styleWorker = new StyleWorker();
        $this->cellWorker = new CellWorker($this->styleWorker);
        $this->labelWorker = new LabelWorker($this->cellWorker);
        $this->tableWorker = new TableWorker($this->cellWorker, $this->labelWorker);
    }

    public function getTableWorker()
    {
        return $this->styleWorker;
    }

    public function getLabelWorker()
    {
        return $this->labelWorker;
    }

    public function getStyleWorker()
    {
        return $this->styleWorker;
    }

    public function getCellWorker()
    {
        return $this->cellWorker;
    }
}