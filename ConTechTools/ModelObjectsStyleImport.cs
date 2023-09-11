using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using WinSys = System.Windows;

namespace ConTechTools
{
    [Transaction(TransactionMode.Manual)]
    public class ModelObjectsStyleImport : IExternalCommand
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
            WinSys.MessageBox.Show("This will be the Object styles Import from CSV botton!\n This Class has not been coded yet.", "MessageBoxTest");

            // Get csv content

            //update or add Object styles

            return Result.Succeeded;
        }

    }
}
