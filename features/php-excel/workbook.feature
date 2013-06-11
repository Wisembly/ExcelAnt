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

Scenario: Apply a default style on the entire Workbook
  Given I create a simple Workbook to be tested
  When I use the writer "Excel5" and I write the Workbook
  Then I should have a "Font" style with the following data on the cell with the coordinates "0,1" in the sheet "0":
    | name    | color  | size |
    | Verdana | FF0000 | 14   |
    And I should have a "Font" style with the following data on the cell with the coordinates "0,1" in the sheet "1":
      | name    | color  | size |
      | Verdana | FF0000 | 14   |

Scenario: We must check if the style of a cell override the default style of the workbook
  Given I create a simple Workbook to be tested
    And I create a StyleCollection and I use it
    And I add a Style "Font" with the following data to the StyleCollection with the index "current":
      | color  | size |
      | 000FFF | 23   |
    And I add a new Cell with the value "override" with the styleCollection with index "current" in the Sheet with index "1" at the coordinates "1,1"
  When I use the writer "Excel5" and I write the Workbook
  Then I should have a "Font" style with the following data on the cell with the coordinates "0,1" in the sheet "0":
    | name    | color  | size |
    | Verdana | FF0000 | 14   |
    And I should have a "Font" style with the following data on the cell with the coordinates "0,1" in the sheet "1":
      | color  | size |
      | 000FFF | 23   |