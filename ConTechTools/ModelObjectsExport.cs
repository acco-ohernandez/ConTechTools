#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

            //Test message
            WinSys.MessageBox.Show("This will be the Object styles Export to CSV botton!", "MessageBoxTest");

            //1: Collect Object style categories into a list
            // List<Categories> allCategories = GetAllCategories(doc);

            //2: Separte categories by type

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

        //private List<Category> GetAllCategories(Document doc)
        //{
        //    //FilteredElementCollector allCategories = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ImportObjectStyles);
        //    FilteredElementCollector collector = new FilteredElementCollector(doc);
        //    collector.OfClass(typeof(Category));
        //    List<Category> allCategories = new List<Category>();
        //    foreach (Element curCategory in collector)
        //    {
        //        if (curCategory.GetType() == typeof(Category))
        //        {
        //            Category c = new Category;
        //            c = curCategory;
        //            allCategories.Add(curCategory);
        //        }
                
        //    }
            
        //    return allCategories;
        //}
    }
}
