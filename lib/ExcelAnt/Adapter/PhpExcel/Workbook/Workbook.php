<?php

namespace ExcelAnt\Adapter\PhpExcel\Workbook;

use \PHPExcel;

use ExcelAnt\Adapter\PhpExcel\Workbook\WorkbookExcelInterface,
    ExcelAnt\Adapter\PhpExcel\Sheet\Sheet,
    ExcelAnt\Sheet\SheetInterface,
    ExcelAnt\Collections\StyleCollection;

class Workbook implements WorkbookExcelInterface
{
    private $phpExcel;
    private $sheetCollection = [];
    private $styleCollection;

    /**
     * @param PHPExcel $phpExcel
     */
    public function __construct(PHPExcel $phpExcel = null)
    {
        $this->phpExcel = $phpExcel ?: new PHPExcel();

        $this->phpExcel->removeSheetByIndex(0);
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
     * {@inheritdoc}
     */
    public function createSheet()
    {
        $sheet = new Sheet($this);
        $this->sheetCollection[] = $sheet;

        return $sheet;
    }

    /**
     * {@inheritdoc}
     */
    public function getSheet($index = 0)
    {
        $this->checkIndexParameter($index);

        return $this->sheetCollection[$index];
    }

    /**
     * {@inheritdoc}
     */
    public function getAllSheets()
    {
        return $this->sheetCollection;
    }

    /**
     * {@inheritdoc}
     */
    public function countSheets()
    {
        return count($this->sheetCollection);
    }

    /**
     * {@inheritdoc}
     */
    public function addSheet(SheetInterface $sheet, $index = null, $insert = false)
    {
        if (isset($index)) {
            if (false === filter_var($index, FILTER_VALIDATE_INT)) {
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
     * {@inheritdoc}
     */
    public function removeSheet($index)
    {
        $this->checkIndexParameter($index);

        unset($this->sheetCollection[$index]);
        $this->sheetCollection = array_values($this->sheetCollection);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setTitle($title)
    {
        $this->phpExcel->getProperties()->setTitle($title);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getTitle()
    {
        return $this->phpExcel->getProperties()->getTitle();
    }

    /**
     * {@inheritdoc}
     */
    public function setCreator($creator)
    {
        $this->phpExcel->getProperties()->setCreator($creator);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getCreator()
    {
        return $this->phpExcel->getProperties()->getCreator();
    }

    /**
     * {@inheritdoc}
     */
    public function setDescription($description)
    {
        $this->phpExcel->getProperties()->setDescription($description);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getDescription()
    {
        return $this->phpExcel->getProperties()->getDescription();
    }

    /**
     * {@inheritdoc}
     */
    public function setCompany($company)
    {
        $this->phpExcel->getProperties()->setCompany($company);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getCompany()
    {
        return $this->phpExcel->getProperties()->getCompany();
    }

    /**
     * {@inheritdoc}
     */
    public function setSubject($subject)
    {
        $this->phpExcel->getProperties()->setSubject($subject);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getSubject()
    {
        return $this->phpExcel->getProperties()->getSubject();
    }

    /**
     * {@inheritdoc}
     */
    public function addStyles(StyleCollection $styles)
    {
        $this->styleCollection = $styles;

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function getStyles()
    {
        return $this->styleCollection;
    }

    /**
     * {@inheritdoc}
     */
    public function hasStyles()
    {
        return empty($this->styleCollection) ? false : true;
    }

    /**
     * Check if the index parameter used by many method to handle the sheetCollection answers the requirements
     *
     * @param integer $index Numeric index of the sheetCollection
     *
     * @throws InvalidArgumentException If the index isn't numeric
     * @throws RuntimeException         If the index doesn't exist
     */
    private function checkIndexParameter($index)
    {
        if (false === filter_var($index, FILTER_VALIDATE_INT)) {
            throw new \InvalidArgumentException("The index must be numeric");
        }

        if (!array_key_exists($index, $this->sheetCollection)) {
            throw new \RuntimeException("The index of this sheet doesn't exist");
        }
    }
}
