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

    public function setTitle($title);

    public function getTitle();

    public function setCreator();

    public function getCreator();

    public function setDescription();

    public function getDescription();

    public function setCompany();

    public function getCompany();

    public function setSubject();

    public function getSubject();

    public function setSecurity();

    public function getSecurity();

    public function setStyle();
}
