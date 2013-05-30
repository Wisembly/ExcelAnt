<?php

namespace ExcelAnt\PhpExcel;

use PHPExcel_Worksheet;

use ExcelAnt\Sheet\SheetInterface;
use ExcelAnt\Table\TableInterface;
use ExcelAnt\Coordinate\Coordinate;
use ExcelAnt\Cell\CellInterface;

class Sheet implements SheetInterface
{
    private $phpExcelWorksheet;
    private $tables = [];
    private $cells = [];

    /**
     * @param WorksheetInterface $worksheet         The parent Worksheet where the Sheet will live
     * @param PHPExcel_Worksheet $phpExcelWorksheet
     */
    public function __construct(Worksheet $worksheet, PHPExcel_Worksheet $phpExcelWorksheet = null)
    {
        $this->phpExcelWorksheet = $phpExcelWorksheet ?: new PHPExcel_Worksheet($worksheet->getRawClass());
    }

    /**
     * Get the raw class
     *
     * @return PHPExcel_Worksheet
     */
    public function getRawClass()
    {
        return $this->phpExcelWorksheet;
    }

    /**
     * {@inheritdoc}
     */
    public function setTitle($name)
    {
        $this->phpExcelWorksheet->setTitle($name);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getTitle()
    {
        return $this->phpExcelWorksheet->getTitle();
    }

    public function addCell(CellInterface $cell, Coordinate $coordinate)
    {
        $this->cells[] = $cell;

        return $this;
    }

    public function getCells()
    {
        return $this->cells;
    }

    /**
     * {@inheritdoc}
     */
    public function addTable(TableInterface $table, Coordinate $coordinate)
    {
        $this->tables[] = $table;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getTables()
    {
        return $this->tables;
    }

    /**
     * {@inheritdoc}
     */
    public function setRowHeight($height, $index)
    {
        if (false === filter_var($height, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->phpExcelWorksheet->getRowDimension($index)->setRowHeight($height);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getRowHeight($index)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        return $this->phpExcelWorksheet->getRowDimension($index)->getRowHeight();
    }

    /**
     * {@inheritdoc}
     */
    public function setColumnWidth($width, $index)
    {
        if (false === filter_var($width, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->phpExcelWorksheet->getColumnDimension($index)->setWidth($width);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getColumnWidth($index)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        return $this->phpExcelWorksheet->getColumnDimension($index)->getWidth();
    }

    public function importImage()
    {

    }
}