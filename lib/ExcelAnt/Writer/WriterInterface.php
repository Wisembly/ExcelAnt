<?php

namespace ExcelAnt\Writer;

use ExcelAnt\Workbook\WorkbookInterface;

interface WriterInterface
{
    public function write(WorkbookInterface $workbook, $path);
}