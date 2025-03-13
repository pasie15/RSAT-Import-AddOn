# RSAT Excel Import Add-on for Tricentis Tosca

This add-on provides functionality to import test cases and test steps from RSAT (Regression Suite Automation Tool) Excel sheets directly into Tricentis Tosca as manual test cases and test steps.

## Features

- Imports test case details and descriptions from the "General" worksheet.
- Imports test step details (actions and expected results) from the "TestCaseSteps" worksheet.
- Automatically creates a manual test case with corresponding manual test steps.
- Easy integration and usage within Tosca Commander.

## Prerequisites

- Tricentis Tosca Commander installed.
- Microsoft Excel installed on the machine (Excel Interop is used).
- Microsoft Excel Interop COM reference added to the project.

## Installation

1. Clone or download the repository.
2. Open the solution in Visual Studio.
3. Ensure the project targets the .NET Framework 4.8).
4. Add reference to Microsoft Excel Interop (`Microsoft Excel XX.X Object Library`) via COM references in your project.
5. Build the solution to generate the DLL.
6. Place the generated DLL into the Tosca Add-ons directory.

## Usage

1. Launch Tosca Commander.
2. Right-click a folder where you want to import the test case.
3. Select "Import RSAT Excel" from the context menu.
4. Ensure your RSAT Excel file (`RSAT_Sample_Test.xlsx`) is placed within the "My Documents" folder.

## Excel Structure

Your RSAT Excel file must include:

- **General Worksheet**: 
  - Cell `B2`: Test Case Name
  - Cell `B3`: Test Case Description

- **TestCaseSteps Worksheet**:
  - Column `B`: Step Action
  - Column `D`: Step Expected Result

## Error Handling

The add-on includes basic error handling with clear messaging for the following scenarios:

- Excel fails to start.
- Errors encountered during Excel operations.
- Any exception during import operations will be displayed as an error message in Tosca.

## Support

For any issues or support, please open an issue in this repository.

## License

This project is provided "as-is" without warranty of any kind.
