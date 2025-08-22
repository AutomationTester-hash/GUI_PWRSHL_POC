
# C# WinAppDriver + Appium + SpecFlow Automation Framework

## Overview
This framework automates running PowerShell commands on a remote Oracle VM using WinAppDriver and Appium, with BDD-style tests written in SpecFlow. It uses the Page Object Model, ExtentReports for reporting, and supports environment/config management.

## Structure
- `Features/` - SpecFlow feature files (BDD scenarios)
- `StepDefinitions/` - Step definition classes for SpecFlow
- `Pages/` - Page Object classes for UI automation
- `Utils/` - Utilities (config, reporting, helpers)
- `Reports/` - ExtentReports output
- `.github/` - Copilot instructions

## Setup
1. Ensure WinAppDriver is running on the VM (default port 4723).
2. Open firewall for port 4723 on the VM.
3. Update `appsettings.json` with VM IP, credentials, and PowerShell path.
4. Build the project:
	```
	dotnet build
	```
5. Run BDD tests:
	```
	dotnet test
	```

## Dependencies
- Appium.WebDriver (4.4.5)
- Selenium.WebDriver (3.141.0)
- SpecFlow & SpecFlow.MsTest
- NUnit (for assertions)
- ExtentReports (5.x)
- Microsoft.Extensions.Configuration & .Json (6.0.0)
- Newtonsoft.Json

## Usage
- Write BDD scenarios in `.feature` files under `Features/`.
- Implement step definitions in `StepDefinitions/`.
- Tests connect to the VM via Appium, launch PowerShell, send commands, and capture output.
- Reports are generated in the `Reports/` folder as HTML.

## Notes
- WinAppDriver must be running and accessible from your local machine.
- Replace config placeholders as needed for your environment.
- Nullable reference warnings are present but do not affect functionality.
