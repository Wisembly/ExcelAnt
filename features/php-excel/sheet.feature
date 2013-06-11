Feature: Test about Sheet

Scenario: Assert there is one sheet in my workbook
  Given I create a Workbook
    And I create a Sheet and I use it
    And I set the following properties to my Sheet with the index "current":
      | title |
      | Foo   |
    And I add all Sheet into the Workbook
  When I use the writer "Excel5" and I write the Workbook
  Then I should see "1" sheet(s)

Scenario: Add two Table in one sheet
  Given I create a Workbook
    And I create a Sheet and I use it
    And I create a Table and I use it
    And I insert the following rows in the Table "current" at the index "null" with the styleCollection "null":
      | rows            |
      | foo1,bar1,baz1  |
      | foofoo1,barbar1 |
    And I insert the Table with the index "current" with the coodinates "1,1" in the Sheet with the index "current"
    And I create a Table and I use it
    And I insert the following rows in the Table "current" at the index "null" with the styleCollection "null":
      | rows            |
      | foo2,bar2,baz2  |
      | foofoo2,barbar2 |
    And I insert the Table with the index "current" with the coodinates "5,1" in the Sheet with the index "current"
    And I add all Sheet into the Workbook
  When I use the writer "Excel5" and I write the Workbook
  Then I should have the value "foo1" in the cell "0,1" of the sheet "0"
    And I should have the value "bar1" in the cell "1,1" of the sheet "0"
    And I should have the value "baz1" in the cell "2,1" of the sheet "0"
    And I should have the value "foofoo1" in the cell "0,2" of the sheet "0"
    And I should have the value "barbar1" in the cell "1,2" of the sheet "0"
    And I should have the value "foo2" in the cell "4,1" of the sheet "0"
    And I should have the value "bar2" in the cell "5,1" of the sheet "0"
    And I should have the value "baz2" in the cell "6,1" of the sheet "0"
    And I should have the value "foofoo2" in the cell "4,2" of the sheet "0"
    And I should have the value "barbar2" in the cell "5,2" of the sheet "0"

Scenario: Test with a top label
  Given I create a Workbook
    And I create a Sheet and I use it
    And I create a Table and I use it
    And I insert the following rows in the Table "current" at the index "null" with the styleCollection "null":
      | rows            |
      | foo,bar,baz  |
      | foofoo,barbar |
    And I set a "top" label of the Table "current" with the following values and with the styleCollection "null":
      | labels |
      | label1 |
      | label2 |
      | label3 |
    And I insert the Table with the index "current" with the coodinates "1,1" in the Sheet with the index "current"
    And I add all Sheet into the Workbook
  When I use the writer "Excel5" and I write the Workbook
  Then I should have the value "label1" in the cell "0,1" of the sheet "0"
    And I should have the value "label2" in the cell "1,1" of the sheet "0"
    And I should have the value "label3" in the cell "2,1" of the sheet "0"
    And I should have the value "foo" in the cell "0,2" of the sheet "0"
    And I should have the value "bar" in the cell "1,2" of the sheet "0"
    And I should have the value "baz" in the cell "2,2" of the sheet "0"
    And I should have the value "foofoo" in the cell "0,3" of the sheet "0"
    And I should have the value "barbar" in the cell "1,3" of the sheet "0"

Scenario: Test with a left label
  Given I create a Workbook
    And I create a Sheet and I use it
    And I create a Table and I use it
    And I insert the following rows in the Table "current" at the index "null" with the styleCollection "null":
      | rows            |
      | foo,bar,baz  |
      | foofoo,barbar |
    And I set a "left" label of the Table "current" with the following values and with the styleCollection "null":
      | labels |
      | label1 |
      | label2 |
    And I insert the Table with the index "current" with the coodinates "1,1" in the Sheet with the index "current"
    And I add all Sheet into the Workbook
  When I use the writer "Excel5" and I write the Workbook
  Then I should have the value "label1" in the cell "0,1" of the sheet "0"
    And I should have the value "label2" in the cell "0,2" of the sheet "0"
    And I should have the value "foo" in the cell "1,1" of the sheet "0"
    And I should have the value "bar" in the cell "2,1" of the sheet "0"
    And I should have the value "baz" in the cell "3,1" of the sheet "0"
    And I should have the value "foofoo" in the cell "1,2" of the sheet "0"
    And I should have the value "barbar" in the cell "2,2" of the sheet "0"

Scenario: Test with a full label
  Given I create a Workbook
    And I create a Sheet and I use it
    And I create a Table and I use it
    And I insert the following rows in the Table "current" at the index "null" with the styleCollection "null":
      | rows            |
      | foo,bar,baz  |
      | foofoo,barbar |
    And I set a "full" label of the Table "current" with the following values and with the styleCollection "null":
      | top    | left   |
      | label1 | label4 |
      | label2 | label5 |
      | label3 |        |
    And I insert the Table with the index "current" with the coodinates "1,1" in the Sheet with the index "current"
    And I add all Sheet into the Workbook
  When I use the writer "Excel5" and I write the Workbook
  Then I should have the value "label1" in the cell "1,1" of the sheet "0"
    And I should have the value "label2" in the cell "2,1" of the sheet "0"
    And I should have the value "label3" in the cell "3,1" of the sheet "0"
    And I should have the value "label4" in the cell "0,2" of the sheet "0"
    And I should have the value "label5" in the cell "0,3" of the sheet "0"
    And I should have the value "foo" in the cell "1,2" of the sheet "0"
    And I should have the value "bar" in the cell "2,2" of the sheet "0"
    And I should have the value "baz" in the cell "3,2" of the sheet "0"
    And I should have the value "foofoo" in the cell "1,3" of the sheet "0"
    And I should have the value "barbar" in the cell "2,3" of the sheet "0"
