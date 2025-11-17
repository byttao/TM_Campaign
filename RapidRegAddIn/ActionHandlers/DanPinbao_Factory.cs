using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using RapidRegAddIn.Utilities;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace RapidRegAddIn.ActionHandlers
{
    public class DanPinbao_Factory
    {
        private string _folderPath = "";

        public void CreateParams(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException($"“{nameof(path)}”不能为 null 或空白。", nameof(path));
            }

            _folderPath = path;

            Globals.ThisAddIn.Application.ScreenUpdating = false;
            try
            {
                CreateProductSKUFile();
            }
            catch (Exception e)
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                MessageBox.Show(e.ToString());
                throw;
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        /// <summary>
        /// 删除空白行
        /// </summary>
        /// <param name="sh">要处理工作表</param>
        /// <param name="columnName">要处理的列名</param>
        private void FilterAndDeleteRowsWithBlankValue(Worksheet sh, string columnName)
        {
            Range usedRange = sh.UsedRange;
            Range columnToFilter = null;
            Range productFilter = null;
            int productColIndex = -1;
            // 查找指定的列
            foreach (Range cell in usedRange.Rows[1].Cells)
            {
                if (cell.Value2.ToString() == columnName)
                {
                    columnToFilter = sh.Columns[cell.Column];
                }
                else if (cell.Value2.ToString() == "商品名称")
                {
                    productColIndex = cell.Column;
                }
            }

            if (productColIndex != -1 && columnToFilter != null)
            {
                // 清除筛选
                sh.AutoFilterMode = false;
                //开启筛选
                sh.Rows["1:1"].AutoFilter();

                try
                {
                    // 开启筛选，并筛选出“商品已删除”
                    sh.Rows[1].AutoFilter(productColIndex, "商品已删除");
                    // 删除筛选结果（可见行，除表头）
                    Range visibleRows = usedRange.Offset[1, 0].Resize[usedRange.Rows.Count - 1].SpecialCells(XlCellType.xlCellTypeVisible);
                    visibleRows.EntireRow.Delete();
                }
                catch (Exception)
                {
                }

                // 关闭筛选
                sh.AutoFilterMode = false;
                //开启筛选
                sh.Rows["1:1"].AutoFilter();

                //分类转为数值
                columnToFilter.TextToColumns(Comma: false, ConsecutiveDelimiter: false, DataType: XlTextParsingType.xlDelimited, Destination: columnToFilter, Other: false,
                    Semicolon: false, Space: false, Tab: false, TextQualifier: XlTextQualifier.xlTextQualifierNone);
                // 关闭筛选
                //_sh.AutoFilterMode = false;
            }
            else
            {
                // 如果未找到指定的列，则显示错误消息
                MessageBox.Show("未找到列：" + columnName);
            }
        }

        private void CreatePivotTable(Workbook wb, string dataSourceSheetName, string pivotTableSheetName)
        {
            Worksheet sh = wb.Sheets[dataSourceSheetName];
            // 定义数据透视表的数据源范围
            Range sourceData = sh.Range["J:O"];

            // 添加一个新的工作表用于放置透视表
            Worksheet pivotSheet = wb.Worksheets.Add();
            pivotSheet.Name = pivotTableSheetName;

            // 定义透视表的目标位置
            Range pivotDestination = pivotSheet.Cells[1, 1];

            // 创建 PivotTable
            PivotTable pivotTable = pivotSheet.PivotTableWizard(XlPivotTableSourceType.xlDatabase,
                sourceData,
                pivotDestination,
                "PivotTable");

            // 设置透视表字段

            pivotTable.PivotFields("商品ID").Orientation = XlPivotFieldOrientation.xlRowField;
            pivotTable.AddDataField(pivotTable.PivotFields("优惠后价格"), "最小值项:优惠后价格", XlConsolidationFunction.xlMin);
            pivotTable.AddDataField(pivotTable.PivotFields("优惠后价格"), "最大值项:优惠后价格", XlConsolidationFunction.xlMax);
            pivotTable.DataPivotField.Orientation = XlPivotFieldOrientation.xlColumnField;
            pivotTable.PivotFields("商品ID").Position = 1;
            pivotTable.DataFields["最小值项:优惠后价格"].NumberFormat = "#,##0.00";
            pivotTable.DataFields["最大值项:优惠后价格"].NumberFormat = "#,##0.00";

            // 激活新的工作表
            pivotSheet.Activate();
        }

        /// <summary>
        /// 单品宝处理
        /// </summary>
        private void CreateProductSKUFile()
        {
            Application excelApp = Globals.ThisAddIn.Application;

            DirectoryInfo directoryInfo = new DirectoryInfo(_folderPath + "\\单品宝");
            FileInfo[] files = directoryInfo.GetFiles("单品宝_*.xlsx");

            if (files.Length > 0)
            {
                Workbook productLevelWb = excelApp.Workbooks.Add();
                Workbook skuLevelWb = excelApp.Workbooks.Add();
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                foreach (FileInfo file in files)
                {
                    Workbook originalWb = excelApp.Workbooks.Open(file.FullName);
                    Worksheet originalSh = originalWb.ActiveSheet;
                    originalWb.Activate();
                    //判断是否为新版单品宝
                    // Variable name suggestion for determining if the sheet is an old version of 单品宝
                    var isOldVersion = originalSh.Range["D1"].Value2.ToString() == "优惠类型";

                    if (!isOldVersion)
                    {
                        // 在D列插入一列
                        originalSh.Columns["D"].Insert();
                        // 在D1单元格写入"优惠类型"
                        originalSh.Range["D1"].Value2 = "优惠类型";
                    }

                    //判断是商品级还是SKU级
                    var masterWorkbook = originalSh.Range["E2"].Value2.ToString() == "商品" ? productLevelWb : skuLevelWb;
                    Worksheet existingSheet = masterWorkbook.Sheets.Cast<Worksheet>().FirstOrDefault(s => s.Name == originalSh.Name);

                    // 如果不存在，则复制工作表到主工作簿
                    if (existingSheet == null)
                    {
                        originalSh.Copy(After: masterWorkbook.Sheets[masterWorkbook.Sheets.Count]);
                    }
                    else
                    {
                        // 如果存在，则将数据从当前工作表复制到已存在的工作表中
                        Range sourceRange = originalSh.UsedRange;
                        sourceRange = sourceRange.Offset[1, 0].Resize[sourceRange.Rows.Count - 1, sourceRange.Columns.Count];
                        Range destinationRange = existingSheet.Cells[existingSheet.UsedRange.Rows.Count + 1, 1];
                        sourceRange.Copy(destinationRange);
                        
                        Foundation.ReleaseObject(sourceRange);
                    }
                    
                    Foundation.ReleaseObject(originalSh);
                    //originalWb.Close(SaveChanges:false);
                    originalWb.Close(false);
                    Foundation.ReleaseObject(originalWb);
                }

                if (productLevelWb.Sheets.Count == 1)
                {
                    Worksheet sh = productLevelWb.Sheets.Add();
                    sh.Name = "0";
                }
                else
                {
                    FilterAndDeleteRowsWithBlankValue(productLevelWb.Sheets["0"], "优惠后价格");
                }

                if (skuLevelWb.Sheets.Count == 1)
                {
                    Worksheet sh = skuLevelWb.Sheets.Add();
                    sh.Name = "0";
                }
                else
                {
                    FilterAndDeleteRowsWithBlankValue(skuLevelWb.Sheets["0"], "优惠后价格");

                    Worksheet sh = skuLevelWb.Sheets["0"];
                    int iRow = sh.Range["A1000000"].End[XlDirection.xlUp].Row;
                    sh.Range[$"Q2:Q{iRow}"].FormulaR1C1 = Foundation.FillExcelWithJSONRules(_folderPath, "SKU级单品宝");

                    CreatePivotTable(skuLevelWb, "0", "Sheet3");
                }

                productLevelWb.SaveAs(_folderPath + "\\单品宝\\单品宝-商品级.xlsx", XlFileFormat.xlOpenXMLWorkbook);
                skuLevelWb.SaveAs(_folderPath + "\\单品宝\\单品宝-SKU级.xlsx", XlFileFormat.xlOpenXMLWorkbook);
                Globals.ThisAddIn.Application.DisplayAlerts = true;
            }
        }
    }
}
