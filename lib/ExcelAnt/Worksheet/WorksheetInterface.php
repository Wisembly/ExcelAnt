<?php

namespace ExcelAnt\Worksheet;

use ExcelAnt\Sheet\SheetInterface;

interface WorksheetInterface
{
    public function getRawClass();

    public function createSheet();

    public function getSheet();

    public function getAllSheets();

    public function countSheets();

    public function addSheet(SheetInterface $sheet, $index = null, $insert = false);

    public function removeSheet($index);

    public function setProperties();

    public function getProperties();

    public function setSecurity();

    public function getSecurity();

    public function setStyle();
}
