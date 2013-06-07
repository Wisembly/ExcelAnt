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
        $this->useContext('style', new StyleContext);
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
        $authorizedProperties = ['title', 'creator', 'description', 'company', 'subject'];

        foreach ($table->getHash() as $values) {

            foreach ($values as $property => $value) {

                if (!in_array($property, $$authorizedProperties)) {
                    continue;
                }

                $method = 'set' . ucfirst($property);

                if (method_exists($this->workbook, $method)) {
                    $this->workbook->$method($value);
                }
            }
        }
    }

    /**
     * @Given /^I add the StyleCollection with the index "([^"]*)" into my Workbook$/
     */
    public function iAddTheStylecollectionWithTheIndexIntoMyWorkbook($index)
    {
        $styleCollection = 'current' === $index ? $this->getSubcontext('style')->styleCollection[$this->getSubcontext('style')->currentStyleCollection] : $this->getSubcontext('style')->styleCollection[$index];
        $this->workbook->addStyles($styleCollection);
    }
}
