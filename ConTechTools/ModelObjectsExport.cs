#region Namespaces
using RApiApp =  Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
//using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using WinSys =  System.Windows;
using SysDraw = System.Drawing;
using System.IO;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
//using System.Windows.Forms;

#endregion

namespace ConTechTools
{
    [Transaction(TransactionMode.Manual)]
    public class ModelObjectsExport : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            RApiApp.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            //***************
            //Test message
            WinSys.MessageBox.Show("This will Export the Model Object Style Settings!", "Exporting MOSS",WinSys.MessageBoxButton.OK,WinSys.MessageBoxImage.Exclamation);
            
            //***************

            // Get the current date and time
            DateTime currentTime = DateTime.Now;

            // Format the date and time as a string
            string dateTimeString = currentTime.ToString("yyyy-MM-dd_HH-mm-ss");

            // Create file path and file name for the Export
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelFileName = Path.Combine(desktopPath, "Model_OSS_Export_" + dateTimeString + ".xlsx");


            List<string> ObjStylesSettingString = new List<string>();
            string header = "Category:LW_Projection:LW_Cut:LineColor:LinePattern:Material";
            Debug.Print(header);

            // get all categories
            Categories categories = doc.Settings.Categories;

            // Sort the parrent cattegories
            List<Category> CatsSortedList = new List<Category>();
            CatsSortedList = SortCategories(categories);

            foreach (Category c in CatsSortedList)
                {
                //if (c.CategoryType == CategoryType.Model || c.CategoryType == CategoryType.Annotation)
                if (c.CategoryType == CategoryType.Model && c.CanAddSubcategory == true)
                {
                    // Output the visible categories
                    if (c.IsVisibleInUI)
                    {
                        //if (c.Name.ToString() == "Lines")
                        //{ Debug.Print("Lines is here ==============================="); }

                        string GetParrentCategoryValues = c.Name + ":" +
                                   (c.GetLineWeight(GraphicsStyleType.Projection)).ToString() + ":" +
                                   (c.GetLineWeight(GraphicsStyleType.Cut)).ToString() + ":" +
                                   GetLineColorName(c.LineColor) + ":" +
                                   GetLineProjLinePatter(doc, c) + ":" +
                                   GetCategoryMaterial(c.Material);

                        Debug.Print(GetParrentCategoryValues);
                        ObjStylesSettingString.Add(GetParrentCategoryValues);

                        CategoryNameMap subCats = c.SubCategories;
                        List<Category> subCatsSortedList = new List<Category>();
                        subCatsSortedList = SortCategories(subCats);
                        if (subCatsSortedList != null)
                        {
                            foreach (Category cat in subCatsSortedList)
                            {
                                string returnedCategory = OutputCatInfo(doc, cat);
                                Debug.Print(returnedCategory);
                                ObjStylesSettingString.Add(returnedCategory);
                            }
                        }
                    }
                }
            }



            // Insert the headers
            ObjStylesSettingString.Insert(0, header);

            // Create the excel file and add the ObjectStylesSettinsg data.
            AddToExcel(excelFileName, ObjStylesSettingString); 

            return Result.Succeeded;
        }

        private List<Category> SortCategories(CategoryNameMap subCats)
        {
            var modelCatsList = new Dictionary<string, Category>();

            foreach (Category c in subCats)
            {
                modelCatsList.Add(c.Name, c);
            }
            var sortedKeys = modelCatsList.Keys.ToList();
            sortedKeys.Sort();
            List<Category> modelCategories = new List<Category>();
            foreach (var key in sortedKeys)
            {
                Debug.Print(modelCatsList[key].Name);
                modelCategories.Add(modelCatsList[key]);
            }
            return modelCategories;
        }

        private string GetCategoryMaterial(Material catMaterial)
        {
            if (catMaterial != null)
                return catMaterial.Name.ToString();
            else
                return "";
        }

        // This method will return the Line Patter Value
        private string GetLineProjLinePatter(Document doc, Category cat)
        {
            ElementId ElemId = cat.GetLinePatternId(GraphicsStyleType.Projection);
            Element catLinePatter = doc.GetElement(ElemId);
            if(catLinePatter != null)
                return catLinePatter.Name;
            else if (ElemId.IntegerValue == -3000010)
                return "Solid";
            else return "";
        }

        private string GetLineColorName(Color lineColor)
        {

            return lineColor.Red.ToString() + '-' +
                   lineColor.Green.ToString() + '-' +
                   lineColor.Blue.ToString();
        }

        private string OutputCatInfo(Document doc, Category cat)
        {
            string rowData;
            // This will de Data to the following columns of the Excel file.
            // CategoryName,LW_Projection,LW_Cut,LineColor,LinePattern,Material

            // ParrentCategory,SubCategoryName Columns
            if (cat.Parent != null)
                rowData = "---| " + cat.Name;//rowData = cat.Parent.Name + ":" + cat.Name;  // Child category name
            else
                rowData = cat.Name + ":------------------------------------:" + cat.Name; // This line should never ouput

            int? projLW = cat.GetLineWeight(GraphicsStyleType.Projection);
            int? cutLW = cat.GetLineWeight(GraphicsStyleType.Cut);

            ElementId projLinePatternId = cat.GetLinePatternId(GraphicsStyleType.Projection);
            ElementId cutLinePatternId = cat.GetLinePatternId(GraphicsStyleType.Cut);

            Element projLineType = doc.GetElement(projLinePatternId);
            Element cutLineType = doc.GetElement(cutLinePatternId);


            // LW_Projection Column
            if (projLW != null)
                rowData += ":" + projLW.ToString();
            else
                rowData += ":";

            // LW_Cut Column
            if (cutLW != null)
                rowData += ":" + cutLW.ToString();
            else
                rowData += ":";

            // Line Color Column
            rowData += ":" + GetLineColorName(cat.LineColor);

            // LinePattern Column
            if (cutLineType != null)
                rowData += ":" + cutLineType.Name;
            else if (projLineType != null)
                rowData += ":" + projLineType.Name;
            else
                rowData += ":Solid";

            // Material Column
            Material mat = cat.Material;
            if (mat != null)
                rowData += ":" + mat.Name.ToString();
            else
                rowData += ":";

            //Debug.Print(rowData);
            return rowData;
        }

        private void AddToExcel(string fileName, List<string> catStrings)
        {
            // Create a new Excel application and workbook
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();

            // Get the first worksheet
            Excel.Worksheet worksheet = workbook.Worksheets.Add();

            // Loop through the string array and output each element to a cell
            for (int i = 0; i < catStrings.Count; i++)
            {
                string[] values = catStrings[i].Split(':');
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
