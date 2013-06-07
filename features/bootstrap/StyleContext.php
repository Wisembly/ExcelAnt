<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use ExcelAnt\Collections\StyleCollection,
    ExcelAnt\Style\Font,
    ExcelAnt\Style\Fill,
    ExcelAnt\Style\Borders,
    ExcelAnt\Style\Border,
    ExcelAnt\Style\Alignment,
    ExcelAnt\Style\Format;

require_once 'PHPUnit/Autoload.php';
require_once 'PHPUnit/Framework/Assert/Functions.php';

/**
 * Style context.
 */
class StyleContext extends BehatContext
{
    private $styleCollection = [];
    private $currentStyleCollection = 0;

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
        error_log(var_export($this->currentStyleCollection, true));
    }

    /**
     * @Given /^I add a Style "([^"]*)" with the following data to the StyleCollection with the index "":$/
     */
    public function iAddAStyleWithTheFollowingDataToTheStylecollectionWithTheIndex($style, TableNode $data)
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
    }
}