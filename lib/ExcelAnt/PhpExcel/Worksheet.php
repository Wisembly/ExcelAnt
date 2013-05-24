<?php

namespace ExcelAnt\PhpExcel;

use PHPExcel;

use ExcelAnt\Worksheet\WorksheetInterface;
use ExcelAnt\Sheet\SheetInterface;
use ExcelAnt\PhpExcel\Sheet;

class Worksheet implements WorksheetInterface
{
    private $phpExcel;
    private $sheetCollection;

    public function __construct(PHPExcel $phpExcel = null)
    {
        $this->phpExcel = $phpExcel ?: new PHPExcel();

        $this->phpExcel->removeSheetByIndex(0);
        $this->sheetCollection = [];
    }

    public function getRawClass()
    {
        return $this->phpExcel;
    }

    public function createSheet()
    {
        $sheet = new Sheet($this);
        $this->sheetCollection[] = $sheet;

        return $sheet;
    }

    public function getSheet($index = 0)
    {
        if (!is_numeric($index)) {
            throw new \InvalidArgumentException("The index must be numeric");
        }

        if (!array_key_exists($index, $this->sheetCollection)) {
            throw new \RuntimeException("The index of this sheet doesn't exist");
        }

        return $this->sheetCollection[$index];
    }

    public function getAllSheets()
    {
        return $this->sheetCollection;
    }

    public function countSheets()
    {
        return count($this->sheetCollection);
    }

    public function addSheet(SheetInterface $sheet, $index = null, $insert = false)
    {
        if (isset($index) && false === $insert) {
            $this->sheetCollection[$index] = $sheet;

            return $this;
        }

        if (isset($index) && true === $insert) {
            $array1 = array_slice($this->sheetCollection, 0, $index);
            $array1[] = $sheet;
            $array2 = array_slice($this->sheetCollection, $index);
            $this->sheetCollection = array_merge($array1, $array2);

            return $this;
        }

        $this->sheetCollection[] = $sheet;

        return $this;
    }

    public function removeSheet()
    {

    }

    public function setProperties()
    {

    }

    public function getProperties()
    {

    }

    public function setSecurity()
    {

    }

    public function getSecurity()
    {

    }

    public function setStyle()
    {

    }
}
