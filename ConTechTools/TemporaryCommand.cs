using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using WinSys = System.Windows;

namespace ConTechTools
{
    [Transaction(TransactionMode.Manual)]
    public class TemporaryCommand : IExternalCommand
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
            WinSys.MessageBox.Show("This is a temporary place holder command class!", "MessageBoxTest");


            //Console.WriteLine("Console.WriteLine - Test===================================");
            //Debug.WriteLine("Debug.WriteLine - Test===================================");
            //Debug.Print("Debug.Print - Test===================================");

            return Result.Succeeded;
        }

    }
}
