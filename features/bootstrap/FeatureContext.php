<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use ExcelAnt\PhpExcel\Workbook,
    ExcelAnt\PhpExcel\Sheet;

//
// Require 3rd-party libraries here:
//
//   require_once 'PHPUnit/Autoload.php';
//   require_once 'PHPUnit/Framework/Assert/Functions.php';
//

/**
 * Features context.
 */
class FeatureContext extends BehatContext
{
    private $workbook;
    private $sheet;

    /**
     * Initializes context.
     * Every scenario gets it's own context object.
     *
     * @param array $parameters context parameters (set them up through behat.yml)
     */
    public function __construct(array $parameters)
    {
        // Initialize your context here
    }

    /**
     * @Given /^I create a Workbook$/
     */
    public function iCreateAWorkbook()
    {
        $this->workbook = new Workbook();
    }

    /**
     * @Given /^I create a Sheet$/
     */
    public function iCreateASheet()
    {
        $this->sheet = new Sheet($this->workbook);
    }

    /**
     * @Then /^I add a sheet in the Workbook$/
     */
    public function iAddASheetInTheWorkbook()
    {
        $this->workbook->addSheet($this->sheet);
        die(var_dump($this->workbook));
    }
}
