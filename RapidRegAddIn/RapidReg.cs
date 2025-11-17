using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using RapidRegAddIn.ActionHandlers;

namespace RapidRegAddIn
{
    public partial class RapidReg
    {
        private string txtFolderName = "";
        private string txtFolderPath = "";

        private void RapidReg_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btn_Danpinbao_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult result = MessageBox.Show("是否确定执行【单品宝生成】功能？", "确认", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                
                Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                bool exists = wb.Sheets.Cast<Worksheet>().Any(x => x.Name == "已报商品列表");
                bool isRegWorkbook = false;
                if (exists)
                {
                    if (wb.Worksheets["已报商品列表"].Range["A1"].Value == "基础信息")
                    {
                        isRegWorkbook = true;
                    }
                }

                if (isRegWorkbook)
                {
                    txtFolderPath = Path.GetDirectoryName(wb.FullName);
                }

                if (txtFolderPath != "")
                {
                    DanPinbao_Factory.CreateParams(txtFolderPath);
                }
                
            }
        }

        private void btn_PathExplorer_Click(object sender, RibbonControlEventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "请选择活动文件夹"; // 设置对话框标题
                DialogResult result = folderDialog.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
                {
                    string selectedFolder = folderDialog.SelectedPath;
                    string folderName = new DirectoryInfo(selectedFolder).Name;

                    // 将路径和文件夹名称显示在文本框中
                    txtFolderPath = selectedFolder;
                    txtFolderName = folderName;
                    l_FolderPath.Label = folderName;
                }
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void btn_Price_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult result = MessageBox.Show("是否确定执行【活动价计算】功能？", "确认", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                // 创建 Stopwatch 对象并开始计时
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                RegPrice_Factory.CreateParams("0");
                // 结束计时并获取经过的时间
                stopwatch.Stop();
                TimeSpan elapsedTime = stopwatch.Elapsed;

                MessageBox.Show("【活动价计算】执行时间为：" + elapsedTime);
            }
        }

        private void btn_Quick_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult result = MessageBox.Show("是否确定执行【表格快速处理】功能？", "确认", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                QuickAction.CreateParams("0");
            }
        }
    }
}
