<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use ExcelAnt\Table\Table;

require_once 'PHPUnit/Autoload.php';
require_once 'PHPUnit/Framework/Assert/Functions.php';

/**
 * Table context.
 */
class TableContext extends BehatContext
{
    public $currentTableIndex = 0;
    public $tableCollection = [];

    public function __construct()
    {

    }

    /**
     * @Given /^I create a Table$/
     */
    public function iCreateATable()
    {
        $this->tableCollection[] = new Table();
    }

    /**
     * @Given /^I create a Table and I use it$/
     */
    public function iCreateATableAndIUseIt()
    {
        $this->iCreateATable();

        $keys = array_keys($this->tableCollection);
        end($keys);

        $this->currentTableIndex = current($keys);
    }

    /**
     * @Given /^I insert the following rows in the Table "([^"]*)" at the index "([^"]*)" with the style "([^"]*)":$/
     */
    public function iInsertTheFollowingRowsInTheTableAtTheIndexWithTheStyle($tableIndex, $rowIndex, $styleIndex, TableNode $rows)
    {
        $tableIndex = 'current' === $tableIndex ? $this->currentTableIndex : $tableIndex;
        $table = $this->getTable($tableIndex);

        foreach ($rows->getHash() as $row) {
            $values = explode(',', $row['rows']);

            foreach ($values as $key => $value) {
                $value = trim($value);
                if (empty($value)) {
                    $values[$key] = null;
                }
            }

            $table->setRow($values, 'null' === $rowIndex ? null : $rowIndex, 'null' === $styleIndex ? null : $styleIndex);
        }
    }

    /**
     * @Given /^I insert the following columns in the Table "([^"]*)" at the index "([^"]*)" with the style "([^"]*)":$/
     */
    public function iInsertTheFollowingColumnsInTheTableAtTheIndexWithTheStyle($tableIndex, $columnIndex, $styleIndex, TableNode $columns)
    {
        $tableIndex = 'current' === $tableIndex ? $this->currentTableIndex : $tableIndex;
        $table = $this->getTable($tableIndex);

        foreach ($columns->getHash() as $column) {
            $values = explode(',', $column['columns']);

            foreach ($values as $key => $value) {
                $value = trim($value);
                if (empty($value)) {
                    $values[$key] = null;
                }
            }

            $table->setColumn($values, 'null' === $columnIndex ? null : $columnIndex, 'null' === $styleIndex ? null : $styleIndex);
        }
    }

    private function getTable($index)
    {
        return $this->tableCollection[$index];
    }
}