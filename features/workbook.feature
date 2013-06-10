Feature: Test about Workbook

Scenario: Assert there is one sheet in my workbook
  Given I create a Workbook
    And I create a Sheet and I use it
    And I set the following properties to my Sheet with the index "current":
      | title |
      | Foo   |
    And I add all Sheet into the Workbook
  When I use the writer "Excel5" and I write the Workbook
  Then I should see "1" sheet(s)