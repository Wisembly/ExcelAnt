<?php

namespace ExcelAnt\Adapter\PhpExcel\Sheet;

use \PHPExcel_Worksheet;

use ExcelAnt\Adapter\PhpExcel\Workbook\Workbook,
    ExcelAnt\Adapter\PhpExcel\Sheet\SheetExcelInterface,
    ExcelAnt\Table\TableInterface,
    ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Cell\CellInterface;

class Sheet implements SheetExcelInterface
{
    private $phpExcelWorksheet;
    private $tables = [];
    private $cells = [];

    /**
     * @param WorkbookInterface $workbook         The parent Workbook where the Sheet will live
     * @param PHPExcel_Worksheet $phpExcelWorksheet
     */
    public function __construct(Workbook $workbook, PHPExcel_Worksheet $phpExcelWorksheet = null)
    {
        $this->phpExcelWorksheet = $phpExcelWorksheet ?: new PHPExcel_Worksheet($workbook->getRawClass());
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

    /**
     * {@inheritdoc}
     */
    public function addCell(CellInterface $cell, Coordinate $coordinate)
    {
        $cell->setCoordinate($coordinate);
        $this->cells[] = $cell;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getCells()
    {
        return $this->cells;
    }

    /**
     * {@inheritdoc}
     */
    public function addTable(TableInterface $table, Coordinate $coordinate)
    {
        $table->setCoordinate($coordinate);
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
            throw new \InvalidArgumentException("Height must be numeric");
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
    public function setDefaultRowHeight($height)
    {
        if (false === filter_var($height, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("Index must be numeric");
        }

        $this->phpExcelWorksheet->getDefaultRowDimension()->setRowHeight($height);

        return $this;
    }

    public function getDefaultRowHeight()
    {
        return $this->phpExcelWorksheet->getDefaultRowDimension()->getRowHeight();
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

        $this->phpExcelWorksheet->getColumnDimensionByColumn($index)->setWidth($width);

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

        return $this->phpExcelWorksheet->getColumnDimensionByColumn($index)->getWidth();
    }

    public function importImage()
    {
        // TODO
    }
}