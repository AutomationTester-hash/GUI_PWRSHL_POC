

# PowerShell GUI Automation Framework (C#, WinAppDriver, Appium, SpecFlow)

## Summary
This framework automates PowerShell command execution on a remote Windows VM using WinAppDriver and Appium, with BDD tests written in SpecFlow. It follows the Page Object Model and generates HTML reports.

## Project Structure
- `Features/`: BDD feature files and code-behind (e.g., `AutomationExamples.feature`)
- `StepDefinitions/`: Step definitions for SpecFlow (e.g., `PowerShellSteps.cs`)
- `Pages/`: Page Object classes for UI automation (e.g., `PowerShellPage.cs`)
- `Utils/`: Utilities for config, driver, PowerShell runner, and reporting (e.g., `ConfigLoader.cs`)

## Getting Started
1. Start WinAppDriver on the target VM (default port 4723).
2. Ensure port 4723 is open on the VM firewall.
3. Update `appsettings.json` with VM details and credentials.
4. Build the project:
	```
	dotnet build
	```
5. Run tests:
	```
	dotnet test --logger "console;verbosity=normal"
	```

	This command runs all tests and outputs results to the console with normal verbosity.


## Main Dependencies (with Purpose)
- **Appium.WebDriver**: For automating Windows application UI via Appium/WinAppDriver
- **Selenium.WebDriver**: Required by Appium for WebDriver protocol support
- **SpecFlow & SpecFlow.MsTest**: For BDD test writing and execution with MsTest runner
- **ExtentReports**: To generate HTML test execution reports
- **Microsoft.Extensions.Configuration & .Json**: For loading and managing configuration from JSON files
- **Newtonsoft.Json**: For JSON parsing and serialization in utilities and config

## Usage Notes
- Write scenarios in `Features/`, implement steps in `StepDefinitions/`.
- Tests connect to the VM, launch PowerShell, send commands, and capture output.
- Reports are generated as HTML in the `Reports/` folder.
- WinAppDriver must be running and accessible.
- Update config as needed for your environment.
