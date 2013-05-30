<?php

namespace ExcelAnt\Worksheet;

use ExcelAnt\Sheet\SheetInterface;
use ExcelAnt\Style\StyleCollection;

interface WorksheetInterface
{
    public function getRawClass();

    /**
     * Create Sheet
     *
     * @return Sheet
     */
    public function createSheet();

    /**
     * Get Sheet
     *
     * @param  integer $index The index of the sheet we want to get
     *
     * @throws InvalidArgumentException If the index isn't numeric
     * @throws RuntimeException         If the index doesn't exist
     *
     * @return Sheet
     */
    public function getSheet();

    /**
     * Get sheetCollection
     *
     * @return array
     */
    public function getAllSheets();

    /**
     * Count sheetCollection
     *
     * @return integer
     */
    public function countSheets();

    /**
     * Add Sheet in sheetCollection
     *
     * You can give an index if you want insert the sheet somewhere. In which case the previous data at this index will be erased.
     *
     * If you give an index and set true the insert parameter, you will insert Sheet in sheetCollection.
     *
     * @param SheetInterface $sheet  The Sheet will be added
     * @param integer        $index  Numeric index of the sheetCollection
     * @param boolean        $insert If you want insert the Sheet or not
     *
     * @throws InvalidArgumentException If the index isn't numeric
     *
     * @return WorksheetInterface
     */
    public function addSheet(SheetInterface $sheet, $index = null, $insert = false);

    /**
     * Remove a sheet
     *
     * @param  integer $index Numeric index of the sheetCollection
     *
     * @throws InvalidArgumentException If the index isn't numeric
     * @throws RuntimeException         If the index doesn't exist
     *
     * @return WorksheetInterface
     */
    public function removeSheet($index);

    /**
     * Set the title of the Worksheet
     *
     * @param mixed $title The title
     */
    public function setTitle($title);

    /**
     * Return the title of the worksheet
     *
     * @return mixed
     */
    public function getTitle();

    /**
     * Set the creator of the worksheet
     *
     * @param mixed $creator
     *
     * @return WorksheetInterface
     */
    public function setCreator($creator);

    /**
     * Get the creator of the Worksheet
     *
     * @return mixed
     */
    public function getCreator();

    /**
     * Set the description of the worksheet
     *
     * @param mixed $description
     *
     * @return WorksheetInterface
     */
    public function setDescription($description);

    /**
     * Get the description of the Worksheet
     *
     * @return mixed
     */
    public function getDescription();

    /**
     * Set the company of the worksheet
     *
     * @param mixed $company
     *
     * @return WorksheetInterface
     */
    public function setCompany($company);

    /**
     * Get the company of the Worksheet
     *
     * @return mixed
     */
    public function getCompany();

    /**
     * Set the subject of the worksheet
     *
     * @param mixed $subject
     *
     * @return WorksheetInterface
     */
    public function setSubject($subject);

    /**
     * Get the subject of the Worksheet
     *
     * @return mixed
     */
    public function getSubject();

    /**
     * Add style
     *
     * @param StyleCollection $styles
     */
    public function addStyles(StyleCollection $styles);

    /**
     * Get styleCollection
     *
     * @return StyleCollection
     */
    public function getStyles();
}
