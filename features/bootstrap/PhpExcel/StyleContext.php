<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use \PHPUnit_Framework_Assert as Assert;

use ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Font,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Borders,
    ExcelAnt\Style\Border,
    ExcelAnt\Style\Alignment,
    ExcelAnt\Style\Format;

/**
 * Style context.
 */
class StyleContext extends BehatContext
{
    public $styleCollection = [];
    public $currentStyleCollection = 0;

    /**
     * @Given /^I create a StyleCollection$/
     */
    public function iCreateAStyleCollection()
    {
        $this->styleCollection[] = new StyleCollection([]);
    }

    /**
     * @Given /^I create a StyleCollection and I use it$/
     */
    public function iCreateAStyleCollectionAndIUseIt()
    {
        $this->iCreateAStyleCollection();

        $keys = array_keys($this->styleCollection);
        end($keys);

        $this->currentStyleCollection = current($keys);
    }

    /**
     * @Given /^I add a Style "([^"]*)" with the following data to the StyleCollection with the index "([^"]*)":$/
     */
    public function iAddAStyle($style, $styleCollectionIndex, TableNode $data)
    {
        switch ($style) {
            case 'Alignment':
                $class = new Alignment();
                break;
            case 'Fill':
                $class = new Fill();
                break;
            case 'Font':
                $class = new Font();
                break;
            case 'Format':
                $class = new Format();
                break;
            default:
                throw new Exception("You must specify a following style : Font, Fill, Alignment, Format");
                break;
        }

        foreach ($data->getHash()[0] as $property => $value) {
            $method = 'set' . ucfirst($property);

            if (method_exists($class, $method)) {
                $class->$method($value);
            }
        }

        $this->styleCollection['current' === $styleCollectionIndex ? $this->currentStyleCollection : $styleCollecitonIndex]->add($class);
    }

    /**
     * @Then /^I should have a "([^"]*)" style with the following data on the cell with the coordinates "(\d+),(\d+)" in the sheet "(\d+)":$/
     */
    public function iShouldHaveAStyleWithTheFollowingDataOnTheCell($style, $x, $y, $sheetIndex, TableNode $data)
    {
        $phpExcelStyle = $this->getMainContext()->excelOutput->getSheet($sheetIndex)->getStyleByColumnAndRow($x, $y);

        switch ($style) {
            case 'Alignment':
                $style = $phpExcelStyle->getAlignment();
                break;
            case 'Fill':
                $style = $phpExcelStyle->getFill();
                break;
            case 'Font':
                $style = $phpExcelStyle->getFont();
                break;
            default:
                throw new Exception("You must specify a following style : Font, Fill, Alignment, Format");
                break;
        }

        foreach ($data->getHash()[0] as $property => $value) {
            $method = 'get' . ucfirst($property);

            if (method_exists($style, $method)) {
                $rawValue = $style->$method();

                if (is_object($rawValue)) {
                    Assert::assertEquals($value, $style->$method()->getRGB());

                    continue;
                }

                Assert::assertEquals($value, $style->$method());
            }
        }
    }
}