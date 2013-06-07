<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use ExcelAnt\PhpExcel\Sheet;

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
}