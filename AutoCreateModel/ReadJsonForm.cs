using Autodesk.Revit.UI;
using System;
using System.Windows.Forms;
using Form = System.Windows.Forms.Form;

namespace AutoCreateModel
{
    public partial class ReadJsonForm : Form
    {
        ExternalEvent externalEvent_CreateToilet;
        //public static string folderPath = @"C:\Prj\Revit\AutoCreateModel\AutoCreateModel\json"; // Json路徑
        public static string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);//桌面路徑
        public ReadJsonForm(UIApplication uiapp, RevitDocument m_connect, IExternalEventHandler handler_CreateToilet)
        {
            InitializeComponent();
            externalEvent_CreateToilet = ExternalEvent.Create(handler_CreateToilet);
            CenterToScreen(); // 置中
        }
        // 選擇Json資料夾路徑
        private void chooseFolderBtn_Click(object sender, EventArgs e)
        {
            // 彈跳視窗選擇Json資料夾
            FolderBrowserDialog path = new FolderBrowserDialog();
            //path.SelectedPath = @"C:\Prj\Revit\AutoCreateModel\AutoCreateModel\json";
            path.ShowDialog();
            this.label1.Text = path.SelectedPath;
            folderPath = path.SelectedPath;
        }
        // 確定
        private void sureBtn_Click(object sender, EventArgs e)
        {
            externalEvent_CreateToilet.Raise(); // 建立廁所
            Close();
        }
        // 取消
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
