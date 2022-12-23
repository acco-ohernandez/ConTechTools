#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
//using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using WinSys =  System.Windows;

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
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            //***************
            //Test message
            //////WinSys.MessageBox.Show("This will be the Object styles Export to CSV botton!", "MessageBoxTest");
            //Console.WriteLine("Console.WriteLine - Test===================================");
            //Debug.WriteLine("Debug.WriteLine - Test===================================");
            //Debug.Print("Debug.Print - Test===================================");
            //***************

            // get all categories
            Categories categories = doc.Settings.Categories;

            int n = categories.Size;

            foreach (Category c in categories)
            {
                if (c.CategoryType == CategoryType.Model || c.CategoryType == CategoryType.Annotation)
                {
                    Debug.Print("---------------------");
                    OutputCatInfo(doc, c);

                    CategoryNameMap subCats = c.SubCategories;

                    if (subCats != null)
                    {
                        foreach (Category cat in subCats)
                        {
                            OutputCatInfo(doc, cat);
                        }
                    }
                }
            }


            //3: export to CSV






            /*
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Transaction Name");
                tx.Commit();
            }
            */

            return Result.Succeeded;
        }

        private void OutputCatInfo(Document doc, Category cat)
        {
            if (cat.Parent != null)
                Debug.Print("    * " + cat.Parent.Name + ", " + cat.Name);
            else
                Debug.Print(cat.CategoryType + " - " + cat.Name);

            int? projLW = cat.GetLineWeight(GraphicsStyleType.Projection);
            int? cutLW = cat.GetLineWeight(GraphicsStyleType.Cut);

            ElementId projLinePatternId = cat.GetLinePatternId(GraphicsStyleType.Projection);
            ElementId cutLinePatternId = cat.GetLinePatternId(GraphicsStyleType.Cut);

            Element projLineType = doc.GetElement(projLinePatternId);
            Element cutLineType = doc.GetElement(cutLinePatternId);

            if (cutLW != null)
                Debug.Print("   Cut LW = " + cutLW.ToString());

            if (projLW != null)
                Debug.Print("   Projection LW = " + projLW.ToString());

            if (cutLineType != null)
                Debug.Print("   Cut Line type = " + cutLineType.Name);

            if (projLineType != null)
                Debug.Print("   Projection Line type = " + projLineType.Name);

            Material mat = cat.Material;

            if (mat != null)
                Debug.Print("   " + mat.Name);

            Color color = cat.LineColor;

            if (color != null)
                Debug.Print("   " + color.Red.ToString()
                    + "," + color.Green.ToString() + ","
                    + color.Blue.ToString());
        }
    }
}
