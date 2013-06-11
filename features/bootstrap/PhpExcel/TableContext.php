<?php

use Behat\Behat\Context\ClosuredContextInterface,
    Behat\Behat\Context\TranslatedContextInterface,
    Behat\Behat\Context\BehatContext,
    Behat\Behat\Exception\PendingException;
use Behat\Gherkin\Node\PyStringNode,
    Behat\Gherkin\Node\TableNode;

use \PHPUnit_Framework_Assert as Assert;

use ExcelAnt\Table\Table,
    ExcelAnt\Table\Label;

/**
 * Table context.
 */
class TableContext extends BehatContext
{
    public $currentTableIndex = 0;
    public $tableCollection = [];

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
     * @Given /^I insert the following rows in the Table "([^"]*)" at the index "([^"]*)" with the styleCollection "([^"]*)":$/
     */
    public function iInsertTheFollowingRowsInTheTableAtTheIndexWithTheStyle($tableIndex, $rowIndex, $styleCollectionIndex, TableNode $rows)
    {
        $table = $this->getTable('current' === $tableIndex ? $this->currentTableIndex : $tableIndex);
        $styleCollection = $this->getStyleCollection($styleCollectionIndex);

        foreach ($rows->getHash() as $row) {
            $values = explode(',', $row['rows']);

            foreach ($values as $key => $value) {
                $value = trim($value);
                if (empty($value)) {
                    $values[$key] = null;
                }
            }

            $table->setRow($values, 'null' === $rowIndex ? null : $rowIndex, $styleCollection);
        }
    }

    /**
     * @Given /^I insert the following columns in the Table "([^"]*)" at the index "([^"]*)" with the styleCollection "([^"]*)":$/
     */
    public function iInsertTheFollowingColumnsInTheTableAtTheIndexWithTheStyle($tableIndex, $columnIndex, $styleCollectionIndex, TableNode $columns)
    {
        $table = $this->getTable('current' === $tableIndex ? $this->currentTableIndex : $tableIndex);
        $styleCollection = $this->getStyleCollection($styleCollectionIndex);

        foreach ($columns->getHash() as $column) {
            $values = explode(',', $column['columns']);

            foreach ($values as $key => $value) {
                $value = trim($value);
                if (empty($value)) {
                    $values[$key] = null;
                }
            }

            $table->setColumn($values, 'null' === $columnIndex ? null : $columnIndex, $styleCollection);
        }
    }

    /**
     * @Given /^I set a "([^"]*)" label of the Table "([^"]*)" with the following values and with the styleCollection "([^"]*)":$/
     */
    public function iSetTheLabelOfTheTableWithTheFollowingValuesAndWithTheStylecollectionCurrent($type, $tableIndex, $styleCollectionIndex, TableNode $table)
    {
        $values = [];

        foreach ($table->getHash() as $value) {

            // Full case
            if ('full' === $type) {
                foreach ($value as $property => $row) {
                    if (!empty($row)) {
                        if ('top' === $property) {
                            $values[0][] = $row;
                        } elseif ('left' === $property) {
                            $values[1][] = $row;
                        }
                    }
                }

                continue;
            }

            $values[] = $value['labels'];
        }

        $styleCollection = $this->getStyleCollection($styleCollectionIndex);

        $label = (new Label($type))->setValues($values, $styleCollection);
        $table = $this->getTable('current' === $tableIndex ? $this->currentTableIndex : $tableIndex);
        $table->setLabel($label);
    }

    private function getTable($index)
    {
        return $this->tableCollection[$index];
    }

    private function getStyleCollection($styleCollectionIndex)
    {
        if ('null' === $styleCollectionIndex) {
            $styleCollection = null;
        } elseif ('current' === $styleCollectionIndex) {
            $styleCollection = $this->getMainContext()->getSubcontext('style')->styleCollection[$this->getMainContext()->getSubcontext('style')->currentStyleCollection];
        } else {
            $styleCollection = $this->getMainContext()->getSubcontext('style')->styleCollection[$styleCollectionIndex];
        }

        return $styleCollection;
    }
}