# RSAT Excel Import Add-on for Tricentis Tosca

This add-on provides functionality to import test cases and test steps from RSAT (Regression Suite Automation Tool) Excel sheets directly into Tricentis Tosca as a test case with folders representing each test step.

## Features

- Imports test case details and descriptions from the "General" worksheet.
- Imports test step details (actions and expected results) from the "TestCaseSteps" worksheet.
- Automatically creates a test case with corresponding manual test steps as folders with test case fields and values concatenated.
- Easy integration and usage within Tosca Commander.

## Prerequisites

- Tricentis Tosca Commander installed.
- Microsoft Excel installed on the machine (Excel Interop is used).
- Microsoft Excel Interop COM reference added to the project.

## Bulding the DLL File

1. Clone or download the repository.
2. Open the solution in Visual Studio.
3. Ensure the project targets the .NET Framework 4.8.
4. Add a reference to Microsoft Excel Interop (`Microsoft Excel XX.X Object Library`) via COM references in your project.
5. Build the solution to generate the RSATImport DLL file.

## Installation

1. Place the RSATImport.dll from the zip file into the C:\Program Files (x86)\TRICENTIS\Tosca Testsuite\ToscaCommander directory.
2. Right-click on the RSATImport.dll file, go to "Properties", and in the "Attributes" section, tick the "Unblock" checkbox.

## Usage

1. Launch Tosca Commander.
2. Right-click a folder where you want to import the test case.
3. Select "RSAT Import AddOn" > "Import RSAT Excel File" from the context menu.
4. Select the RSAT Excel test case file from the File Select Pop-Up Window.

## Excel Structure

Your RSAT Excel file must include:

- **General Worksheet**: 
  - Cell `B4`: Recording Name (Test Case Name)

- **TestCaseSteps Worksheet**:
  - Column `B`: Step Action
  - Column `C`: Field
  - Column `D`: Value

## Error Handling

The add-on includes basic error handling with clear messaging for the following scenarios:

- Excel fails to start.
- Errors encountered during Excel operations.
- Any exception during import or file selection operations will be displayed as an error message in Tosca.

## Reference

Information on how to Create and Install a Tosca AddOn can be found here:
https://documentation.tricentis.com/devcorner/2023.2/tcaddon/topic2.html

Information on classes and properties of Tosca AddOn Objects can be found here:
https://documentation.tricentis.com/devcorner/930/tcaddon/

## Author
**David Taylor**  
[GitHub](https://github.com/pasie15) | [LinkedIn](https://www.linkedin.com/in/david-taylor-96791196/)
