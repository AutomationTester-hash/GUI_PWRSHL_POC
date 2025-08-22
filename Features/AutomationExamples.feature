

 Feature: Automation Examples
  As a user
  I want to automate both GUI and PowerShell scenarios

  Scenario: Run a simple PowerShell command directly
    Given I set the PowerShell execution policy to Unrestricted
    When I run the PowerShell command 'echo Hello from VM'
    Then the PowerShell output should contain 'Hello from VM'

  Scenario: Notepad automation with WinAppDriver
    Given I have initialized the application for automation
  When I enter the command "Hello from Notepad"
  And I erase the text in Notepad after 2 seconds
  
   Scenario: Run a demo executable (Calculator)
    Given I set the PowerShell execution policy to Unrestricted
    When I run the executable 'C:\\Windows\\System32\\calc.exe'
  And I wait for 5 seconds after opening Calculator

  Scenario: Open and close custom app via Automation
    Given I open Oracle VirtualBox Manager
    When I close Oracle VirtualBox Manager
  