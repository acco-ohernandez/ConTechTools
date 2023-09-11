#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
                // Export annotation object styles for each category
                foreach (Category category in annotationCategories)
                {
                    // Export object styles for this category
                    // ExportObjectStyles(doc, category);
                    Debug.Print($"{category.CategoryType} - {category.Name} - {category.IsReadOnly} - {category.LineColor}");
                }
            }
            else
            {
                TaskDialog.Show("No Annotation Categories", "There are no annotation categories in the document.");
            }

            // Return a success result
            return Result.Succeeded;
        }

        // Helper method to get all annotation categories
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

        private List<Category> GetAnnotationCategories2(Document doc)
        {
            Categories categories = doc.Settings.Categories;

            List<Category> annotationCategories = new List<Category>();

            // Use the BuiltInCategory enumeration to filter annotation categories
            foreach (Category cat in categories)
            {
                if (cat.CategoryType == CategoryType.Annotation)
                {
                    annotationCategories.Add(cat);
                }
            }

            // Sort the annotationCategories by category name in ascending order
            //annotationCategories = annotationCategories.OrderBy(cat => cat.Name).ToList();

            return annotationCategories;
        }


        // Helper method to export object styles for a category
        private void ExportObjectStyles(Document doc, Category category)
        {
            // Implement your code to export object styles for the given category here
            // You can use the category variable to identify the category and export styles accordingly
        }


    }
}
