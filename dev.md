

# PowerShell GUI Automation Framework - POC Demo & Developer Guide

## Overview & Use Case
This Proof of Concept (POC) demonstrates how to automate and test Windows GUI applications and PowerShell commands on a remote or local Windows VM. The framework is designed for:
- Automating the launch, interaction, and closure of Windows applications (e.g., VirtualBox, Calculator, Notepad).
- Running and validating PowerShell commands.
- Providing clear, business-readable BDD-style (Given/When/Then) test scenarios.


### POC Scenario & What It Solves
This POC addresses the need to:
- Remotely automate and validate Windows GUI and PowerShell operations for regression, smoke, or integration testing.
- Provide a maintainable, extensible, and business-friendly automation solution.
- Enable non-technical users to understand and run automation with minimal setup.

**Example Scenarios Automated:**
- Open and close Oracle VirtualBox Manager.
- Launch Calculator, wait, and close it.
- Run a custom executable from config.
- Run PowerShell commands and check output.
- Erase text in Notepad.

---

## Tech Stack & Why Chosen
- **C# (.NET 6):** Modern, robust, and widely used for Windows automation.
- **Appium + WinAppDriver:** Industry-standard for automating Windows GUI applications using the WebDriver protocol.
- **SpecFlow (with MsTest):** Enables BDD-style test writing, making scenarios easy to read and maintain.
- **ExtentReports:** Generates professional, interactive HTML reports for test runs.
- **Microsoft.Extensions.Configuration:** Simplifies configuration management via JSON files.
- **Custom Utilities:** For configuration loading, driver management, PowerShell execution, and reporting.

**Why this stack?**
- Supports both technical and non-technical users.
- Easily extensible for new apps or scenarios.
- Open-source and widely supported tools.
- Clean separation of concerns (feature files, step logic, page objects, utilities).

---

## How to Run the Demo
1. Start WinAppDriver on the target VM (default port 4723).
2. Ensure port 4723 is open on the VM firewall.
3. Update `appsettings.json` with VM details, credentials, and app paths.
4. Build the project:
   ```
   dotnet build
   ```
5. Run the tests:
   ```
   dotnet test --logger "console;verbosity=normal"
   ```
6. View the HTML report in the `Reports/` folder.

---

## What Is Included in This POC?
- **Feature Files:** Example BDD scenarios for automating apps and PowerShell.
- **Step Definitions:** C# code that implements each test step, with clear comments.
- **Page Objects:** Encapsulate UI interactions for maintainability.
- **Utility Classes:** Helpers for config, driver management, PowerShell execution, and reporting.
- **Config File:** `appsettings.json` for all environment and app settings.
- **HTML Reports:** (In Progress) The framework is being integrated with ExtentReports to generate professional HTML reports for each test run.

---

## Possible Questions & Answers

**Q: What is the main purpose of this framework?**
A: To automate and test Windows GUI applications and PowerShell commands using BDD scenarios, with easy reporting and configuration.

**Q: What technologies does it use and why?**
A: C#/.NET 6 (robust Windows support), Appium + WinAppDriver (GUI automation), SpecFlow (readable BDD tests), ExtentReports (HTML reporting), Microsoft.Extensions.Configuration (easy config), and custom utilities for maintainability.

**Q: How do I add a new test scenario?**
A: Write a new `.feature` file in the `Features/` folder and implement the steps in `StepDefinitions/PowerShellSteps.cs`.

**Q: How do I change which VM or app is automated?**
A: Update the relevant settings in `appsettings.json` (VM IP, credentials, app paths, etc.).

**Q: Where do I find the test results?**
A: In the `Reports/` folder (HTML report) and `TestResults/` folder (raw results).

**Q: What if an app doesn't launch or a test fails?**
A: Check the paths in `appsettings.json`, ensure WinAppDriver is running, and review the console and HTML report for errors.

**Q: Can I extend this framework?**
A: Yes! Add new feature files, step definitions, or page objects as needed. Utility classes help with config, driver, and reporting.

**Q: Is technical knowledge required to run the demo?**
A: No. Just follow the steps above. For advanced changes, some C# and automation knowledge is helpful.

**Q: How do I explain the architecture to someone else?**
A: The framework separates business logic (feature files), step logic (step definitions), UI automation (page objects), and utilities (config, driver, reporting) for clarity and maintainability.

**Q: What are the main benefits for a business or team?**
A: Easy to demo, hand over, and extend. Non-technical users can understand scenarios. Technical users can add new features quickly.

**Q: What if I want to automate a different app?**
A: Add a new page object for the app, update the config, and write new feature/step definitions as needed.

---

## Quick Reference: Project Structure
- `Features/` - BDD scenarios
- `StepDefinitions/` - Step logic
- `Pages/` - UI automation helpers
- `Utils/` - Config, driver, PowerShell, reporting
- `Reports/` - HTML reports
- `TestResults/` - Raw test results
- `appsettings.json` - All settings

---THE END---
