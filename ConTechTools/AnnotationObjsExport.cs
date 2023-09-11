#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace ConTechTools
{
    [Transaction(TransactionMode.Manual)]
    public class AnnotationObjsExport : IExternalCommand
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
                Application app = uiapp.Application;
                Document doc = uidoc.Document;

                // Show a message to confirm that the command is executing
                TaskDialog.Show("Annotation Object Styles Export", "Command is executing.");

                // Get all annotation categories
                List<Category> annotationCategories = GetAnnotationCategories(doc);

                if (annotationCategories != null && annotationCategories.Any())
                {
                    // Export annotation object styles for each category, including sub-settings
                    foreach (Category category in annotationCategories)
                    {
                        Debug.Print($"{category.Name} - {category.IsReadOnly} - {category.LineColor}");
                        ExportObjectStyles(doc, category);

                        // Export sub-settings for this category
                        ExportSubSettings(doc, category);
                    }
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
            Categories categories = doc.Settings.Categories;

            // Get all categories and sort them by name in ascending order
            List<Category> allCategories = categories.Cast<Category>().OrderBy(cat => cat.Name).ToList();

            List<Category> annotationCategories = new List<Category>();

            // Filter for annotation categories
            foreach (Category cat in allCategories)
            {
                if (cat.CategoryType == CategoryType.Annotation)
                {
                    annotationCategories.Add(cat);
                }
            }

            return annotationCategories;
        }

        // Helper method to export sub-settings of a category
        // ...

        // Helper method to export sub-settings of a category
        private void ExportSubSettings(Document doc, Category category)
        {
            CategoryNameMap subCategories = category.SubCategories;

            if (subCategories != null && subCategories.Size > 0)
            {
                // Sort the subCategories by name
                List<Category> sortedSubCategories = subCategories.Cast<Category>()
                    .OrderBy(subCat => subCat.Name).ToList();

                foreach (Category subCategory in sortedSubCategories)
                {
                    Debug.Print($"|-- {subCategory.Name} - {subCategory.IsReadOnly} - {subCategory.LineColor}");
                    ExportObjectStyles(doc, subCategory);

                    // Recursively export sub-settings of this sub-category
                    ExportSubSettings(doc, subCategory);
                }
            }
        }


        // Helper method to export object styles for a category
        // Helper method to export object styles for a category
        private void ExportObjectStyles(Document doc, Category category)
        {
            // Check if the Desktop folder exists
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (!Directory.Exists(desktopPath))
            {
                Debug.Print("Desktop folder not found.");
                return;
            }

            // Define the CSV file path
            string csvFilePath = Path.Combine(desktopPath, "AnnotationObjectStyles.csv");

            try
            {
                using (StreamWriter writer = new StreamWriter(csvFilePath, true))
                {
                    // Write category information to the CSV file
                    writer.Write($"Category Name: {category.Name},");
                    writer.Write($"Is Read-Only: {category.IsReadOnly},");
                    writer.Write($"Line Color: {category.LineColor}");
                    writer.WriteLine(); // Add an empty line for separation

                    // Implement your code here to export specific object styles for the given category
                    // Example: You can iterate through the object styles and write them to the CSV file.
                    // Replace the following code with your actual implementation.

                    // writer.WriteLine("Object Style 1: Some Details");
                    // writer.WriteLine("Object Style 2: Some Details");
                    // ...

                    // End of object styles

                    //writer.WriteLine(); // Add an empty line for separation
                }

                Debug.Print($"Object styles for '{category.Name}' exported to '{csvFilePath}'.");
            }
            catch (Exception ex)
            {
                Debug.Print($"Error exporting object styles for '{category.Name}': {ex.Message}");
            }
        }

    }
}
