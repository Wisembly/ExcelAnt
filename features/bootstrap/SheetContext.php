<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use \PHPUnit_Framework_Assert as Assert;

use ExcelAnt\Adapter\PhpExcel\Sheet\Sheet,
    ExcelAnt\Coordinate\Coordinate,
    ExcelAnt\Cell\Cell,
    ExcelAnt\Cell\EmptyCell;

require_once 'PHPUnit/Autoload.php';
require_once 'PHPUnit/Framework/Assert/Functions.php';

/**
 * Sheet context.
 */
class SheetContext extends BehatContext
{
    public $currentSheetIndex = 0;
    public $sheetCollection = [];

    public function __construct()
    {
        $this->useContext('table', new TableContext);
    }

    /**
     * @Given /^I create a Sheet$/
     */
    public function iCreateASheet()
    {
        $this->sheetCollection[] = new Sheet($this->getMainContext()->workbook);
    }

    /**
     * @Given /^I create a Sheet and I use it$/
     */
    public function iCreateASheetAndIUseIt()
    {
        $this->iCreateASheet();

        $keys = array_keys($this->sheetCollection);
        end($keys);

        $this->currentSheetIndex = current($keys);
    }

    /**
     * @Given /^I set the following properties to my Sheet with the index "([^"]*)":$/
     */
    public function iSetTheFollowingPropertiesToMySheet($index, TableNode $table)
    {
        $authorizedProperties = ['title'];
        $index = 'current' === $index ? $this->currentSheetIndex : $index;

        foreach ($table->getHash() as $values) {

            foreach ($values as $property => $value) {

                if (!in_array($property, $authorizedProperties)) {
                    continue;
                }

                $method = 'set' . ucfirst($property);

                if (method_exists($this->sheetCollection[$index], $method)) {
                    $this->sheetCollection[$index]->$method($value);
                }
            }
        }
    }

    /**
     * @Given /^I add a new Cell with the value "([^"]*)" with the styleCollection with index "([^"]*)" in the Sheet with index "([^"]*)" at the coordinates "([^"]*)"$/
     */
    public function iAddANewCellWithTheValueWithTheStylecollectionWithIndexInTheSheetWithIndexAtTheCoordinates($value, $styleCollectionIndex, $sheetIndex, $coordinate)
    {
        $coordinate = explode(',', $coordinate);
        $coordinate = new Coordinate($coordinate[0], $coordinate[1]);

        if ('null' === $styleCollectionIndex) {
            $styleCollection = null;
        } elseif ('current' === $styleCollectionIndex) {
            $styleCollection = $this->getMainContext()->getSubcontext('style')->styleCollection[$this->getMainContext()->getSubcontext('style')->currentStyleCollection];
        } else {
            $styleCollection = $this->getMainContext()->getSubcontext('style')->styleCollection[$styleCollectionIndex];
        }

        if ('null' === $value) {
            $cell = new EmptyCell();
        } else {
            $cell = new Cell($value);
        }

        $this->sheetCollection['current' === $sheetIndex ? $this->currentSheetIndex : $sheetIndex]->addCell($cell, $coordinate);
    }

    /**
     * @Given /^I set the row "([^"]*)" height with the value "([^"]*)" of the Sheet with index "([^"]*)"$/
     */
    public function iSetTheRowHeightOfTheSheetWithIndex($value, $indexRow, $index)
    {
        $value = (int) $value;
        $this->sheetCollection['current' === $index ? $this->currentSheetIndex : $index]->setRowHeight($value, $indexRow);
    }

    /**
     * @Given /^I set the column "([^"]*)" width with the value "([^"]*)" of the Sheet with index "([^"]*)"$/
     */
    public function iSetTheColumnWidthWithTheValueOfTheSheetWithIndex($value, $indexColumn, $index)
    {
        $value = (int) $value;
        $this->sheetCollection['current' === $index ? $this->currentSheetIndex : $index]->setColumnWidth($value, $indexColumn);
    }

    /**
     * @Then /^I should see "([^"]*)" sheet\(s\)$/
     */
    public function iShouldSeeSheetS($number)
    {
        Assert::assertEquals($number, $this->getMainContext()->excelOutput->getSheetCount());
    }
}