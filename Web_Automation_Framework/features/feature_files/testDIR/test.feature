Feature: summa

  Scenario Outline: test
    Given the URL is opened
    Then Click gender
    Then Enter your full name "<SheetName>"
    Then Enter your Email "<SheetName>"
    Then Enter your phonenumber "<SheetName>"
    Then Scroll until visible and enter name in section 1
    Then close the driver


    Examples:
      | SheetName    |
      | Sample sheet 1|

