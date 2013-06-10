<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException,
    Behat\Behat\Event\ScenarioEvent;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use \PHPUnit_Framework_Assert as Assert;

use ExcelAnt\PhpExcel\Workbook,
    ExcelAnt\PhpExcel\Sheet,
    ExcelAnt\PhpExcel\Writer\PhpExcelWriter\Excel5,
    ExcelAnt\PhpExcel\Writer\Writer,
    ExcelAnt\PhpExcel\Writer\Worker\CellWorker,
    ExcelAnt\PhpExcel\Writer\Worker\LabelWorker,
    ExcelAnt\PhpExcel\Writer\Worker\StyleWorker,
    ExcelAnt\PhpExcel\Writer\Worker\TableWorker;

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

    /**
     * @When /^I use the writer "([^"]*)" and I write the Workbook$/
     */
    public function iUseTheWriterAndIWriteTheWorkbook($arg1)
    {
        $styleWorker = new StyleWorker();
        $cellWorker = new CellWorker($styleWorker);
        $labelWorker = new LabelWorker($cellWorker);
        $tableWorker = new TableWorker($cellWorker, $labelWorker);
        $writer = new Writer(new Excel5('./features/behat.xls'), $tableWorker, $cellWorker, $styleWorker);
        $phpExcel = $writer->convert($this->workbook);
        $writer->write($phpExcel);

        $this->excelOutput = (new PHPExcel_Reader_Excel5())
            ->load('./features/behat.xls');
    }

    /**
     * @Then /^I should see "([^"]*)" sheet\(s\)$/
     */
    public function iShouldSeeSheetS($number)
    {
        Assert::assertEquals($number, $this->excelOutput->getSheetCount());
    }
}
