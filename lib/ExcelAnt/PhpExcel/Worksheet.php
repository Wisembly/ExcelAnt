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

    /**
     * Get the raw class
     *
     * @return PHPExcel
     */
    public function getRawClass()
    {
        return $this->phpExcel;
    }

    /**
     * Create a Sheet
     *
     * @return Sheet
     */
    public function createSheet()
    {
        $sheet = new Sheet($this);
        $this->sheetCollection[] = $sheet;

        return $sheet;
    }

    /**
     * Get a Sheet
     *
     * @param  integer $index The index of the sheet we want to get
     *
     * @throws InvalidArgumentException If the index isn't a numeric value
     * @throws RuntimeException If the index doesn't exist
     *
     * @return Sheet
     */
    public function getSheet($index = 0)
    {
        $this->checkIndexParameter($index);

        return $this->sheetCollection[$index];
    }

    /**
     * Get all the Sheet object of the sheetCollection
     *
     * @return array
     */
    public function getAllSheets()
    {
        return $this->sheetCollection;
    }

    /**
     * Return the sheets number
     *
     * @return integer
     */
    public function countSheets()
    {
        return count($this->sheetCollection);
    }

    /**
     * Add a new sheet in the sheetCollection
     *
     * You can give an index if you want insert the sheet somewhere. In which case the previous data at this index will be erased.
     *
     * If you give an index and set true the insert parameter, you will insert the Sheet in the array.
     *
     * @param SheetInterface $sheet  The sheet will be added
     * @param integer        $index  Numeric index of the sheetCollection
     * @param boolean        $insert If you want insert the sheet or not
     *
     * @throws InvalidArgumentException If the index isn't numeric
     *
     * @return Worksheet
     */
    public function addSheet(SheetInterface $sheet, $index = null, $insert = false)
    {
        if (isset($index)) {
            if (!is_numeric($index)) {
                throw new \InvalidArgumentException("The index must be numeric");
            }

            if (false === $insert) {
                $this->sheetCollection[$index] = $sheet;

                return $this;
            }

            if (true === $insert) {
                $array1 = array_slice($this->sheetCollection, 0, $index);
                $array1[] = $sheet;
                $array2 = array_slice($this->sheetCollection, $index);
                $this->sheetCollection = array_merge($array1, $array2);

                return $this;
            }
        }

        $this->sheetCollection[] = $sheet;

        return $this;
    }

    /**
     * Remove a sheet
     *
     * @param  integer $index Numeric index of the sheetCollection
     *
     * @throws InvalidArgumentException If the index isn't a numeric value
     * @throws RuntimeException If the index doesn't exist
     *
     * @return Worksheet
     */
    public function removeSheet($index)
    {
        $this->checkIndexParameter($index);

        unset($this->sheetCollection[$index]);
        $this->sheetCollection = array_values($this->sheetCollection);

        return $this;
    }

    /**
     * Set the tile of the Worksheet
     *
     * @param mixed $title The title
     */
    public function setTitle($title)
    {
        $this->phpExcel->getProperties()->setTitle($title);

        return $this;
    }

    /**
     * Return the title of the worksheet
     *
     * @return mixed
     */
    public function getTitle()
    {
        return $this->phpExcel->getProperties()->getTitle();
    }

    /**
     * Set the creator of the worksheet
     *
     * @param mixed $creator
     *
     * @return Worksheet
     */
    public function setCreator($creator)
    {
        $this->phpExcel->getProperties()->setCreator($creator);

        return $this;
    }

    /**
     * Get the creator of the Worksheet
     *
     * @return mixed
     */
    public function getCreator()
    {
        return $this->phpExcel->getProperties()->getCreator();
    }

    /**
     * Set the description of the worksheet
     *
     * @param mixed $description
     *
     * @return Worksheet
     */
    public function setDescription($description)
    {
        $this->phpExcel->getProperties()->setDescription($description);

        return $this;
    }

    /**
     * Get the description of the Worksheet
     *
     * @return mixed
     */
    public function getDescription()
    {
        return $this->phpExcel->getProperties()->getDescription();
    }

    /**
     * Set the company of the worksheet
     *
     * @param mixed $company
     *
     * @return Worksheet
     */
    public function setCompany($company)
    {
        $this->phpExcel->getProperties()->setCompany($company);

        return $this;
    }

    /**
     * Get the company of the Worksheet
     *
     * @return mixed
     */
    public function getCompany()
    {
        return $this->phpExcel->getProperties()->getCompany();
    }

    /**
     * Set the subject of the worksheet
     *
     * @param mixed $subject
     *
     * @return Worksheet
     */
    public function setSubject($subject)
    {
        $this->phpExcel->getProperties()->setSubject($subject);

        return $this;
    }

    /**
     * Get the subject of the Worksheet
     *
     * @return mixed
     */
    public function getSubject()
    {
        return $this->phpExcel->getProperties()->getSubject();
    }

    public function setStyle()
    {

    }

    /**
     * Check if the index parameter used by many method to handle the sheetCollection answers the requirements
     *
     * @param integer $index Numeric index of the sheetCollection
     *
     * @throws InvalidArgumentException If the index isn't a numeric value
     * @throws RuntimeException If the index doesn't exist
     */
    private function checkIndexParameter($index)
    {
        if (!is_numeric($index)) {
            throw new \InvalidArgumentException("The index must be numeric");
        }

        if (!array_key_exists($index, $this->sheetCollection)) {
            throw new \RuntimeException("The index of this sheet doesn't exist");
        }
    }
}
