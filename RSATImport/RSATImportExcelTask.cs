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
        public override string Name => "Import RSAT Excel";

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
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\RSAT_Sample_Test.xlsx";

                Workbook workbook = excelApp.Workbooks.Open(filePath);
                Worksheet generalSheet = workbook.Worksheets["General"];
                Worksheet stepsSheet = workbook.Worksheets["TestCaseSteps"];

                // Read general test case details
                string rsatTestCaseName = generalSheet.Cells[2, 2].Value?.ToString();
                string rsatTestCaseDescription = generalSheet.Cells[3, 2].Value?.ToString();

                // Create test case in Tosca
                TestCase toscaTestCase = destFolder.CreateManualTestCase(destFolder);
                toscaTestCase.Name = rsatTestCaseName;
                toscaTestCase.Description = rsatTestCaseDescription;

                // Read test steps from "TestCaseSteps" sheet
                Range stepsRange = stepsSheet.UsedRange;
                object[,] stepsData = (object[,])stepsRange.Value2;
                int stepRows = stepsData.GetLength(0);

                for (int row = 2; row <= stepRows; row++)
                {
                    string rsatStepAction = stepsData[row, 2]?.ToString();
                    string rsatStepValue = stepsData[row, 4]?.ToString();

                    if (!string.IsNullOrWhiteSpace(rsatStepAction))
                    {
                        var toscaTestStep = toscaTestCase.CreateManualXTestStep();
                        toscaTestStep.Name = rsatStepAction;

                        var toscaStepValue = toscaTestStep.CreateManualXTestStepValue();
                        toscaStepValue.Value = rsatStepValue;
                        toscaStepValue.ActionMode = XTestStepActionMode.Verify;
                    }
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(generalSheet);
                Marshal.ReleaseComObject(stepsSheet);
                Marshal.ReleaseComObject(workbook);

                taskContext.ShowMessageBox("Success", "RSAT test case and steps imported successfully.");

                return null;
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
