<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException,
    Behat\Behat\Event\ScenarioEvent;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use \PHPUnit_Framework_Assert as Assert;

use ExcelAnt\Adapter\PhpExcel\Workbook\Workbook,
    ExcelAnt\Adapter\PhpExcel\Sheet\Sheet,
    ExcelAnt\Adapter\PhpExcel\Writer\PhpExcelWriter\Excel5,
    ExcelAnt\Adapter\PhpExcel\Writer\Writer,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\CellWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\LabelWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\StyleWorker,
    ExcelAnt\Adapter\PhpExcel\Writer\Worker\TableWorker,
    ExcelAnt\Coordinate\Coordinate;

/**
 * Features context.
 */
class FeatureContext extends BehatContext
{
    public $workbook;
    public $excelOutput;

    /**
     * @param array $parameters context parameters (set them up through behat.yml)
     */
    public function __construct(array $parameters)
    {
        $this->useContext('sheet', new SheetContext);
        $this->useContext('style', new StyleContext);
    }

    /**
     * @AfterScenario
     */
     public function removeExportFile(ScenarioEvent $event)
     {
        unlink('./features/behat.xls');
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

                if (!in_array($property, $authorizedProperties)) {
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
     * @Given /^I create a simple Workbook to be tested$/
     */
    public function iCreateASimpleWorkbookToBeTested()
    {
        $this->iCreateAWorkbook();
        $this->getSubcontext('sheet')->iCreateASheet();
        $this->getSubcontext('style')->iCreateAStyleCollectionAndIUseIt();

        $styleData = new TableNode();
        $styleData->setRows([['name', 'color', 'size'], ['Verdana', 'ff0000', 14]]);
        $this->getSubcontext('style')->iAddAStyle('Font', 'current', $styleData);

        $this->getSubcontext('sheet')->iAddANewCell('Foo', "null", "current", "1,1");
        $this->getSubcontext('sheet')->iCreateASheetAndIUseIt();
        $this->getSubcontext('sheet')->iAddANewCell('Bar', "null", "current", "1,1");
        $this->iAddAllSheetIntoTheWorkbook();

        $this->iAddTheStylecollectionWithTheIndexIntoMyWorkbook("current");
    }

    /**
     * @Given /^I add the StyleCollection with the index "([^"]*)" into my Workbook$/
     */
    public function iAddTheStylecollectionWithTheIndexIntoMyWorkbook($index)
    {
        $styleCollection = 'current' === $index ? $this->getSubcontext('style')->styleCollection[$this->getSubcontext('style')->currentStyleCollection] : $this->getSubcontext('style')->styleCollection[$index];
        $this->workbook->addStyles($styleCollection);
    }

    /**
     * @When /^I use the writer "([^"]*)" and I write the Workbook$/
     */
    public function iUseTheWriterAndIWriteTheWorkbook($writer)
    {
        $styleWorker = new StyleWorker();
        $cellWorker = new CellWorker($styleWorker);
        $labelWorker = new LabelWorker($cellWorker);
        $tableWorker = new TableWorker($cellWorker, $labelWorker);
        $writer = new Writer(new Excel5('./features/behat.xls'), $tableWorker, $cellWorker, $styleWorker);

        // If there isn't Sheet, there is an issue with the export
        // So, we add one
        if (0 === $this->workbook->countSheets()) {
            $this->workbook->addSheet(new Sheet($this->workbook));
        }

        $phpExcel = $writer->convert($this->workbook);
        $writer->write($phpExcel);

        $this->excelOutput = (new PHPExcel_Reader_Excel5())
            ->load('./features/behat.xls');
    }

    /**
     * @Then /^I should have "([^"]*)" as title of the workbook$/
     */
    public function iShouldHaveAsTitleOfTheWorkbook($title)
    {
        Assert::assertEquals($title, $this->excelOutput->getProperties()->getTitle());
    }

    /**
     * @Then /^I should have "([^"]*)" as creator of the workbook$/
     */
    public function iShouldHaveAsCreatorOfTheWorkbook($creator)
    {
        Assert::assertEquals($creator, $this->excelOutput->getProperties()->getCreator());
    }

    /**
     * @Then /^I should have "([^"]*)" as description of the workbook$/
     */
    public function iShouldHaveAsDescriptionOfTheWorkbook($description)
    {
        Assert::assertEquals($description, $this->excelOutput->getProperties()->getDescription());
    }

    /**
     * @Then /^I should have "([^"]*)" as company of the workbook$/
     */
    public function iShouldHaveAsCompanyOfTheWorkbook($company)
    {
        Assert::assertEquals($company, $this->excelOutput->getProperties()->getCompany());
    }

    /**
     * @Then /^I should have "([^"]*)" as subject of the workbook$/
     */
    public function iShouldHaveAsSubjectOfTheWorkbook($subject)
    {
        Assert::assertEquals($subject, $this->excelOutput->getProperties()->getSubject());
    }
}
