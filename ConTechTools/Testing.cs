using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using OfficeOpenXml; // Add EPPlus namespace

namespace ConTechTools
{
    [Transaction(TransactionMode.Manual)]
    public class Testing : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                var app = uiapp.Application;
                Document doc = uidoc.Document;

                // Show a message to confirm that the command is executing
                //TaskDialog.Show("Annotation Object Styles Export", "Command is executing.");

                // Get all annotation categories
                List<Category> annotationCategories = GetAnnotationCategories(doc);

                if (annotationCategories != null && annotationCategories.Any())
                {
                    // Create a list to store data for exporting to EPPlus
                    List<ExportData> exportDataList = new List<ExportData>();

                    // Export annotation object styles for each category, including sub-settings
                    foreach (Category category in annotationCategories)
                    {
                        Debug.Print($"{category.Name} - {category.IsReadOnly} - {category.LineColor.Red}{category.LineColor.Green}{category.LineColor.Blue}");

                        // Export object styles and add data to the list
                        ExportObjectStyles(doc, category, exportDataList);


                        // Export sub-settings for this category
                        ExportSubSettings(doc, category, exportDataList);
                    }

                    // Export data to EPPlus with Save As functionality
                    ExportDataToEPPlusWithSaveAs(exportDataList);
                }
                else
                {
                    TaskDialog.Show("No Annotation Categories", "There are no annotation categories in the document.");
                }

                // Return a success result
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // Helper method to get all annotation categories, including sub-settings
        private List<Category> GetAnnotationCategories(Document doc)
        {
            // Get the categories from the document settings
            Categories categories = doc.Settings.Categories;

            // Use LINQ to filter and sort the categories
            List<Category> annotationCategories = categories.Cast<Category>()
                // Filter categories based on CategoryType
                .Where(cat => cat.CategoryType == CategoryType.Annotation)
                // Sort categories by name in ascending order
                .OrderBy(cat => cat.Name)
                // Convert the result to a List<Category>
                .ToList();

            // Return the list of annotation categories
            return annotationCategories;
        }

        // Helper method to export object styles for a category
        private void ExportObjectStyles(Document doc, Category category, List<ExportData> exportDataList, string prefix = "")
        {
            // Check if the category is null
            if (category == null)
            {
                Debug.Print("Category is null.");
                return;
            }

            // Create an ExportData object with the required information
            ExportData exportData = new ExportData
            {
                CategoryName = $"{prefix}{category.Name}", // Include the prefix in the category name
                LineWeight = category.GetLineWeight(GraphicsStyleType.Projection)?.ToString() ?? "N/A",
                LinePattern = GetLinePatternName(doc, category),
            };

            // Check if GraphicsStyle is available
            if (category.GetGraphicsStyle(GraphicsStyleType.Projection) is GraphicsStyle graphicsStyle)
            {
                // Check if LineColor is available and not null
                var lineColor = graphicsStyle.GraphicsStyleCategory.LineColor;
                if (lineColor != null)
                {
                    // Create a Revit color from RGB values
                    Autodesk.Revit.DB.Color revitColor = new Autodesk.Revit.DB.Color(
                        lineColor.Red,
                        lineColor.Green,
                        lineColor.Blue
                    );

                    exportData.LineColor = revitColor;
                }
            }

            // Add the ExportData object to the list
            exportDataList.Add(exportData);
        }

        // Helper method to export sub-settings of a category
        private void ExportSubSettings(Document doc, Category category, List<ExportData> exportDataList, string prefix = "")
        {
            // Get the sub-categories of the category
            CategoryNameMap subCategories = category.SubCategories;

            if (subCategories != null && subCategories.Size > 0)
            {
                // Define the prefix for sub-categories
                string subCategoryPrefix = " |--";

                // Iterate through the sub-categories
                foreach (Category subCategory in subCategories)
                {

                    // Print sub-category details (you can remove this if not needed)
                    Debug.Print($"{prefix}{subCategoryPrefix}{subCategory.Name} - {subCategory.IsReadOnly} - {subCategory.LineColor}");

                    // Export object styles and add data to the list
                    ExportObjectStyles(doc, subCategory, exportDataList, prefix + subCategoryPrefix);

                    // Recursively export sub-settings of this sub-category with the updated prefix
                    ExportSubSettings(doc, subCategory, exportDataList, prefix + subCategoryPrefix);
                }
            }
        }


        private string GetLinePatternName(Document doc, Category category)
        {
            // Get the ElementId of the line pattern for the specified category
            ElementId linePatternId = category.GetLinePatternId(GraphicsStyleType.Projection);

            // Try to get the line pattern element from the document using the ElementId
            Element linePattern = doc.GetElement(linePatternId);

            if (linePattern != null)
            {
                // If a line pattern element was found, return its name
                return linePattern.Name;
            }
            else if (linePatternId.IntegerValue == -3000010)
            {
                // Check if the line pattern Id corresponds to a solid line (-3000010)
                return "Solid";
            }
            else
            {
                // If no line pattern element was found and it's not a solid line, return "N/A"
                return "N/A";
            }
        }

        // Helper method to export data to EPPlus with Save As functionality
        private void ExportDataToEPPlusWithSaveAs(List<ExportData> exportDataList)
        {

            // Create a file save dialog
            var saveFileDialog = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Save Exported Data As",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                DefaultExt = "xlsx",
                AddExtension = true
            };

            // Show the file save dialog
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelFilePath = saveFileDialog.FileName;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Create a new Excel package
                using (var package = new ExcelPackage())
                {
                    // Add a new worksheet
                    var worksheet = package.Workbook.Worksheets.Add("AnnotationObjectStyles");

                    // Set headers
                    worksheet.Cells[1, 1].Value = "Category";
                    worksheet.Cells[1, 2].Value = "Line Weight";
                    worksheet.Cells[1, 3].Value = "Line Color (RGB)";
                    worksheet.Cells[1, 4].Value = "Line Pattern";

                    // Fill in the data
                    int row = 2; // Start from the second row
                    foreach (var exportData in exportDataList)
                    {
                        worksheet.Cells[row, 1].Value = exportData.CategoryName;
                        worksheet.Cells[row, 2].Value = exportData.LineWeight;
                        worksheet.Cells[row, 3].Value = $"{exportData.LineColor.Red}-{exportData.LineColor.Green}-{exportData.LineColor.Blue}";
                        worksheet.Cells[row, 4].Value = exportData.LinePattern;
                        row++;
                    }

                    // Save the Excel file to the chosen location
                    File.WriteAllBytes(excelFilePath, package.GetAsByteArray());
                    Debug.Print($"Data exported to '{excelFilePath}'.");
                }
            }
        }

        // Data structure to store export data
        private class ExportData
        {
            public string CategoryName { get; set; }
            public string LineWeight { get; set; }
            public Autodesk.Revit.DB.Color LineColor { get; set; }
            public string LinePattern { get; set; }
        }
    }
}

