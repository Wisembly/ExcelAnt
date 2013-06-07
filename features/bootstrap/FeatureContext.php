<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use ExcelAnt\PhpExcel\Workbook,
    ExcelAnt\PhpExcel\Sheet;

require_once 'PHPUnit/Autoload.php';
require_once 'PHPUnit/Framework/Assert/Functions.php';

/**
 * Features context.
 */
class FeatureContext extends BehatContext
{
    public $workbook;

    /**
     * @param array $parameters context parameters (set them up through behat.yml)
     */
    public function __construct(array $parameters)
    {
        $this->useContext('sheet', new SheetContext);
    }

    /**
     * @Given /^I create a Workbook$/
     */
    public function iCreateAWorkbook()
    {
        $this->workbook = new Workbook();
    }

    /**
     * @Given /^I add the Sheet with the index "([^"]*)" into the Workbook$/
     */
    public function iAddTheSheetWithTheIndexIntoTheWorkbook($index)
    {
        $this->workbook->addSheet($this->getSubContext('sheet')->sheetCollection[$index]);
    }

    /**
     * @Given /^I add all Sheet into the Workbook$/
     */
    public function iAddAllSheetIntoTheWorkbook()
    {
        foreach ($this->getSubContext('sheet')->sheetCollection as $key => $sheet) {
            $this->iAddTheSheetWithTheIndexIntoTheWorkbook($key);
        }
    }

    /**
     * @Given /^I set the following properties to my Workbook:$/
     */
    public function iSetTheFollowingPropertiesToMyWorkbook(TableNode $table)
    {
        foreach ($table->getHash() as $values) {

            foreach ($values as $property => $value) {
                $method = 'set' . ucfirst($property);

                if (method_exists($this->workbook, $method)) {
                    $this->workbook->$method($value);
                }
            }
        }
    }
}
