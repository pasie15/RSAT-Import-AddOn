using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using Tricentis.TCAddOns;
using Tricentis.TCAPIObjects.Objects;

namespace RSATImportAddon.Tasks
{
    public class RSATImportExcelTask : TCAddOnTask
    {
        public override string Name => "Import RSAT Excel File";

        public override Type ApplicableType => typeof(TCFolder);

        public override bool IsTaskPossible(TCObject obj)
        {
            return obj is TCFolder folder && folder.PossibleContent.Split(';').Contains("TestCase");
        }

        public override bool RequiresChangeRights => true;

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
            {
                taskContext.ShowErrorMessage("Excel Error", "Could not start Excel. Ensure Office is installed.");
                return null;
            }

            try
            {
                var destFolder = objectToExecuteOn as TCFolder;
                // Allow user to select Excel file
                string[] selectedFiles = taskContext.GetFilePaths("Select RSAT Excel File", false, "xlsx");
                if (selectedFiles == null || selectedFiles.Length == 0)
                {
                    taskContext.ShowErrorMessage("File Selection Error", "No file selected.");
                    return null;
                }

                string filePath = selectedFiles[0];

                Workbook workbook = excelApp.Workbooks.Open(filePath);
                Worksheet generalSheet = workbook.Worksheets["General"];
                Worksheet stepsSheet = workbook.Worksheets["TestCaseSteps"];

                // Read general test case details
                string rsatTestCaseName = generalSheet.Cells[4, 2].Value?.ToString();

                // Create test case in Tosca
                TestCase toscaTestCase = destFolder.CreateTestCase();
                toscaTestCase.Name = rsatTestCaseName;

                // Create Precondition Folder Structure
                var preconditionFolder = toscaTestCase.CreateFolder();
                preconditionFolder.Name = "Precondition";

                var loginFolder = preconditionFolder.CreateFolder();
                loginFolder.Name = "Log into Dynamics 365 FinOps Environment";

                // Create Process Folder
                var processFolder = toscaTestCase.CreateFolder();
                processFolder.Name = "Process";

                // Read test steps from "TestCaseSteps" sheet
                Range stepsRange = stepsSheet.UsedRange;
                object[,] stepsData = (object[,])stepsRange.Value2;
                int stepRows = stepsData.GetLength(0);

                for (int row = 3; row <= stepRows; row++)
                {
                    string rsatStepAction = stepsData[row, 2]?.ToString();
                    string rsatStepField = stepsData[row, 3]?.ToString();
                    string rsatStepValue = stepsData[row, 4]?.ToString();

                    if (!string.IsNullOrWhiteSpace(rsatStepAction))
                    {
                        string folderName = rsatStepAction;

                        if (!string.IsNullOrWhiteSpace(rsatStepField))
                            folderName += " | " + rsatStepField;

                        if (!string.IsNullOrWhiteSpace(rsatStepValue))
                            folderName += ": " + rsatStepValue;

                        var toscaTestStepFolder = processFolder.CreateFolder();
                        toscaTestStepFolder.Name = folderName;
                    }
                }

                // Create Postcondition Folder
                var postconditionFolder = toscaTestCase.CreateFolder();
                postconditionFolder.Name = "Postcondition";

                workbook.Close(false);
                Marshal.ReleaseComObject(generalSheet);
                Marshal.ReleaseComObject(stepsSheet);
                Marshal.ReleaseComObject(workbook);

                taskContext.ShowMessageBox("Success", "RSAT test case and steps imported successfully.");
                return toscaTestCase;
            }
            catch (Exception ex)
            {
                taskContext.ShowErrorMessage("Import Error", ex.Message);
                return null;
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}