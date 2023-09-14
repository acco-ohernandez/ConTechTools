using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ConTechTools
{
    [Transaction(TransactionMode.Manual)]
    public class ModelObjectsExport_refactored : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                // Retrieve Revit application and document
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Show a message to confirm that the command is executing
                // TaskDialog.Show("Model Object Styles Export", "Command is executing.");

                // Get all model categories and their settings
                List<string> modelObjectStyles = ExportModelObjectStyles(doc);

                // Generate the Excel file and save the data
                SaveModelObjectStylesToExcel(modelObjectStyles);

                // Return a success result
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private List<string> ExportModelObjectStyles(Document doc)
        {
            List<string> modelObjectStyles = new List<string>();

            // Iterate through categories and gather object style information
            foreach (Category category in GetModelCategories(doc))
            {
                string categoryInfo = ProcessCategory(doc, category);
                modelObjectStyles.Add(categoryInfo);
            }

            return modelObjectStyles;
        }

        private IEnumerable<Category> GetModelCategories(Document doc)
        {
            return doc.Settings.Categories
                .Cast<Category>()
                .Where(category =>
                    category.CategoryType == CategoryType.Model &&
                    !CategoryIsInExceptionList(category.Name));
        }

        private bool CategoryIsInExceptionList(string categoryName)
        {
            string[] categoriesExceptionList =
            {
                "Analysis Display Style",
                "Analysis Results",
                // Add more categories as needed
            };

            return categoriesExceptionList.Contains(categoryName);
        }

        private string ProcessCategory(Document doc, Category category)
        {
            // Extract and format category information
            string categoryInfo = $"{category.Name}:{GetLineWeight(doc, category, GraphicsStyleType.Projection)}:" +
                $"{GetLineWeight(doc, category, GraphicsStyleType.Cut)}:" +
                $"{GetColorString(category.LineColor)}:" +
                $"{GetLinePatternName(doc, category)}:" +
                $"{GetCategoryMaterial(category.Material)}";

            return categoryInfo;
        }

        private string GetLineWeight(Document doc, Category category, GraphicsStyleType styleType)
        {
            int? lineWeight = category.GetLineWeight(styleType);
            return lineWeight?.ToString() ?? "";
        }

        private string GetColorString(Color color)
        {
            return $"{color.Red}-{color.Green}-{color.Blue}";
        }

        private string GetLinePatternName(Document doc, Category category)
        {
            ElementId linePatternId = category.GetLinePatternId(GraphicsStyleType.Projection);
            Element linePattern = doc.GetElement(linePatternId);

            if (linePattern != null)
            {
                return linePattern.Name;
            }
            else if (linePatternId.IntegerValue == -3000010)
            {
                return "Solid";
            }
            else
            {
                return "";
            }
        }

        private string GetCategoryMaterial(Material catMaterial)
        {
            return catMaterial?.Name ?? "";
        }

        private void SaveModelObjectStylesToExcel(List<string> modelObjectStyles)
        {
            // Create a new Excel file and add the object style settings
            string excelFilePath = GenerateExcelFilePath();
            AddToExcel(excelFilePath, modelObjectStyles);
        }

        private string GenerateExcelFilePath()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string dateTimeString = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            return Path.Combine(desktopPath, $"Model_OSS_Export_{dateTimeString}.xlsx");
        }

        private void AddToExcel(string fileName, List<string> modelObjectStyles)
        {
            // Create a new Excel application and workbook
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();

            // Get the first worksheet
            Excel.Worksheet worksheet = workbook.Worksheets.Add();

            // Loop through the string array and output each element to a cell
            for (int i = 0; i < modelObjectStyles.Count; i++)
            {
                string[] values = modelObjectStyles[i].Split(':');
                for (int j = 0; j < values.Length; j++)
                {
                    worksheet.Cells[i + 1, j + 1].Value = values[j];
                }
            }

            // Save the workbook and close the Excel application
            workbook.SaveAs(fileName);
            excelApp.Quit();
            Process.Start(fileName);
        }
    }
}
