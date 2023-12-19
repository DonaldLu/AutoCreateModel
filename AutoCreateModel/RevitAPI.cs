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
        //static string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        static string addinAssmeblyPath = @"C:\ProgramData\Autodesk\Revit\Addins\2020\Sino_Station\"; // 封包版路徑位址
        public Result OnStartup(UIControlledApplication a)
        {
            addinAssmeblyPath = addinAssmeblyPath + "AutoCreateModel.dll";

            RibbonPanel ribbonPanel = null;
            try { a.CreateRibbonTab("捷運規範校核"); } catch { }
            try { ribbonPanel = a.CreateRibbonPanel("捷運規範校核", "自動建模"); }
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
            // 在面板上添加一個按鈕, 點擊此按鈕觸動AutoCreateModel.CreateMRT
            PushButton createToiletBtn = ribbonPanel.AddItem(new PushButtonData("AutoCreateModel", "建立廁所", addinAssmeblyPath, "AutoCreateModel.CreateMRT")) as PushButton;
            createToiletBtn.LargeImage = convertFromBitmap(Properties.Resources.Toilet);

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
