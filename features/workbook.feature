Feature: Workbook handler and determining if the output is correctly provided

Scenario: Assert there is one sheet in my workbook
  Given I create a Workbook
    And I set the following properties to my Workbook:
      | title        | creator   | description                   | company  | subject         |
      | Behat Export | Wirsembly | An export wich will be tested | Wisembly | My behat export |
  When I use the writer "Excel5" and I write the Workbook
  Then I should have "Behat Export" as title of the workbook
  Then I should have "Wirsembly" as creator of the workbook
  Then I should have "An export wich will be tested" as description of the workbook
  Then I should have "My behat export" as subject of the workbook