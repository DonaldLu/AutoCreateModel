using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace AutoCreateModel
{
    public class RevitAPI : IExternalApplication
    {
        static string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        //static string checkPlatformPath = @"C:\ProgramData\Autodesk\Revit\Addins\2018\AutoCreateModel\AutoCreateModel.dll";
        public Result OnStartup(UIControlledApplication a)
        {
            RibbonPanel ribbonPanel = null;
            try { a.CreateRibbonTab("自動建模"); } catch { }
            try { ribbonPanel = a.CreateRibbonPanel("自動建模", "建立廁所"); }
            catch
            {
                List<RibbonPanel> panel_list = new List<RibbonPanel>();
                panel_list = a.GetRibbonPanels("自動建模");
                foreach (RibbonPanel rp in panel_list)
                {
                    if (rp.Name == "建立廁所")
                    {
                        ribbonPanel = rp;
                    }
                }
            }

            PushButton pushbutton1 = ribbonPanel.AddItem(new PushButtonData("AutoCreateModel", "建立廁所", addinAssmeblyPath, "AutoCreateModel.CreateMRT")) as PushButton;
            pushbutton1.LargeImage = convertFromBitmap(Properties.Resources.house);

            return Result.Succeeded;
        }

        BitmapSource convertFromBitmap(System.Drawing.Bitmap bitmap)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
