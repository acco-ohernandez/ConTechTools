#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace ConTechTools
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            //Step 1: Create a ribbon tab
            try
            {
                a.CreateRibbonTab("ConTech");
            }
            catch (Exception)
            {
                Debug.Print("TAB ALREADY EXISTS");
            }

            //Step 2: Create a ribbon panel
            //RibbonPanel curPanel = a.CreateRibbonPanel("ConTech", "ConTech Panel"); // this would fail if the panel already exists
            RibbonPanel curPanel = CreateRibbonPanel(a, "ConTech", "ConTech Panel"); // Created my own method to check if panel exists firts

            //Step 3: Create button data instances (pbd = Push Button Data)
            PushButtonData pData1 = new PushButtonData("MOS_Export", "Export\rObject Styles", GetAssemblyName(), "ConTechTools.ModelObjectsExport");
            PushButtonData pData2 = new PushButtonData("MOS_Import", "Import\rObject Styles", GetAssemblyName(), "ConTechTools.ModelObjectsStyleImport");
            PushButtonData pData3 = new PushButtonData("AOS_Export", "Export\rAnnotationStyles", GetAssemblyName(), "ConTechTools.AnnotationObjsExport");
            PushButtonData pData4 = new PushButtonData("Button4Name", "pbNotInUse4", GetAssemblyName(), "ConTechTools.TemporaryCommand");
            PushButtonData pData5 = new PushButtonData("Button5Name", "pbNotInUse5", GetAssemblyName(), "ConTechTools.TemporaryCommand");
            PushButtonData pData6 = new PushButtonData("Button6Name", "pbNotInUse6", GetAssemblyName(), "ConTechTools.TemporaryCommand");
            PushButtonData pData7 = new PushButtonData("Button7Name", "pbNotInUse7", GetAssemblyName(), "ConTechTools.TemporaryCommand");

            SplitButtonData sData1 = new SplitButtonData("splitButton1", "Split Button1");
            PulldownButtonData pbData1 = new PulldownButtonData("pulldownButton1", "Pulldown\rButton1");

            //Step 4: Add images
            pData1.Image = BitmapToImageSource(ConTechTools.Properties.Resources.ModelObjectStyles_MOS_16x16);
            pData1.LargeImage = BitmapToImageSource(ConTechTools.Properties.Resources.ModelObjectStyles_MOS_32x32);

            pData2.Image = BitmapToImageSource(ConTechTools.Properties.Resources.CoolFace_16x16);
            pData2.LargeImage = BitmapToImageSource(ConTechTools.Properties.Resources.CoolFace_32x32);

            pData3.Image = BitmapToImageSource(ConTechTools.Properties.Resources.QuestionMark_16x16);
            pData3.LargeImage = BitmapToImageSource(ConTechTools.Properties.Resources.QuestionMark_32x32);

            pData4.Image = BitmapToImageSource(ConTechTools.Properties.Resources.Test_16x16);
            pData4.LargeImage = BitmapToImageSource(ConTechTools.Properties.Resources.Test_32x32);

            pData5.Image = BitmapToImageSource(ConTechTools.Properties.Resources.CoolFace_16x16);
            pData5.LargeImage = BitmapToImageSource(ConTechTools.Properties.Resources.CoolFace_32x32);

            pData6.Image = BitmapToImageSource(ConTechTools.Properties.Resources.QuestionMark_16x16);
            pData6.LargeImage = BitmapToImageSource(ConTechTools.Properties.Resources.QuestionMark_32x32);

            pData7.Image = BitmapToImageSource(ConTechTools.Properties.Resources.Test_16x16);
            pData7.LargeImage = BitmapToImageSource(ConTechTools.Properties.Resources.Test_32x32);


            pbData1.Image = BitmapToImageSource(ConTechTools.Properties.Resources.Test_16x16);
            pbData1.LargeImage = BitmapToImageSource(ConTechTools.Properties.Resources.Test_32x32);
            //Step 5: Add tooltip info


            //Step 6: Create buttons
            PushButton b1 = curPanel.AddItem(pData1) as PushButton;

            curPanel.AddStackedItems(pData2, pData3);

            SplitButton splitButton1 = curPanel.AddItem(sData1) as SplitButton;
            splitButton1.AddPushButton(pData4);
            splitButton1.AddPushButton(pData5);

            PulldownButton pulldownButton1 = curPanel.AddItem(pbData1) as PulldownButton;
            pulldownButton1.AddPushButton(pData6);
            pulldownButton1.AddPushButton(pData7);

            return Result.Succeeded;
        }

        private BitmapImage BitmapToImageSource(System.Drawing.Bitmap bm)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();
                return bmi;
            }
        }

        // This method will take a png file and load it into memory
        private RibbonPanel CreateRibbonPanel(UIControlledApplication a, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in a.GetRibbonPanels(tabName))
            {
                // if tmpPanel already exists in the tab, return the existing panel
                if (tmpPanel.Name == panelName)
                {
                    return tmpPanel;
                }
            }
            // if the tmpPanel was not found, Create panelName
            RibbonPanel returnPanel = a.CreateRibbonPanel(tabName, panelName);
            return returnPanel;
        }

        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
