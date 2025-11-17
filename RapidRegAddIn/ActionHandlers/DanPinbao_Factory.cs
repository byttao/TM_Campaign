using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using RapidRegAddIn.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace RapidRegAddIn.ActionHandlers
{
    public static class DanPinbao_Factory
    {
        private static string _folderPath = "";

        public static void CreateParams(string path)
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
        private static void FilterAndDeleteRowsWithBlankValue(Excel.Worksheet sh, string columnName)
        {
            Excel.Range usedRange = sh.UsedRange;
            Excel.Range columnToFilter = null;

            // 查找指定的列
            foreach (Excel.Range cell in usedRange.Rows[1].Cells)
            {
                if (cell.Value2.ToString() == columnName)
                {
                    columnToFilter = sh.Columns[cell.Column];
                    break;
                }
            }

            if (columnToFilter != null)
            {
                // 清除筛选
                sh.AutoFilterMode = false;
                //开启筛选
                sh.Rows["1:1"].AutoFilter();
                // 启用筛选并筛选出空白值的行
                columnToFilter.SpecialCells(Excel.XlCellType.xlCellTypeBlanks).EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                //分类转为数值
                columnToFilter.TextToColumns(Comma: false, ConsecutiveDelimiter: false, DataType: Excel.XlTextParsingType.xlDelimited, Destination: columnToFilter, Other: false,
                    Semicolon: false, Space: false, Tab: false, TextQualifier: Excel.XlTextQualifier.xlTextQualifierNone);
                // 关闭筛选
                //_sh.AutoFilterMode = false;
            }
            else
            {
                // 如果未找到指定的列，则显示错误消息
                MessageBox.Show("未找到列：" + columnName);
            }
        }

        private static void CreatePivotTable(Excel.Workbook wb, string dataSourceSheetName, string pivotTableSheetName)
        {
            Excel.Worksheet sh = wb.Sheets[dataSourceSheetName];
            // 定义数据透视表的数据源范围
            Excel.Range sourceData = sh.Range["J:O"];

            // 添加一个新的工作表用于放置透视表
            Excel.Worksheet pivotSheet = wb.Worksheets.Add();
            pivotSheet.Name = pivotTableSheetName;

            // 定义透视表的目标位置
            Excel.Range pivotDestination = pivotSheet.Cells[1, 1];

            // 创建 PivotTable
            Excel.PivotTable pivotTable = pivotSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase,
                sourceData,
                pivotDestination,
                "PivotTable");

            // 设置透视表字段

            pivotTable.PivotFields("商品ID").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.AddDataField(pivotTable.PivotFields("优惠后价格"), "最小值项:优惠后价格", Excel.XlConsolidationFunction.xlMin);
            pivotTable.AddDataField(pivotTable.PivotFields("优惠后价格"), "最大值项:优惠后价格", Excel.XlConsolidationFunction.xlMax);
            pivotTable.DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            pivotTable.PivotFields("商品ID").Position = 1;
            pivotTable.DataFields["最小值项:优惠后价格"].NumberFormat = "#,##0.00";
            pivotTable.DataFields["最大值项:优惠后价格"].NumberFormat = "#,##0.00";

            // 激活新的工作表
            pivotSheet.Activate();
        }

        /// <summary>
        /// 单品宝处理
        /// </summary>
        private static void CreateProductSKUFile()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;

            DirectoryInfo directoryInfo = new DirectoryInfo(_folderPath + "\\单品宝");
            FileInfo[] files = directoryInfo.GetFiles("单品宝_*.xlsx");

            if (files.Length > 0)
            {
                Excel.Workbook productLevelWb = excelApp.Workbooks.Add();
                Excel.Workbook skuLevelWb = excelApp.Workbooks.Add();
                foreach (FileInfo file in files)
                {
                    Excel.Workbook originalWb = excelApp.Workbooks.Open(file.FullName);
                    Excel.Worksheet originalSh = originalWb.ActiveSheet;
                    //判断是商品级还是SKU级
                    var masterWorkbook = originalSh.Range["E2"].Value2.ToString() == "商品" ? productLevelWb : skuLevelWb;
                    Excel.Worksheet existingSheet = masterWorkbook.Sheets.Cast<Excel.Worksheet>().FirstOrDefault(s => s.Name == originalSh.Name);

                    // 如果不存在，则复制工作表到主工作簿
                    if (existingSheet == null)
                    {
                        originalSh.Copy(After: masterWorkbook.Sheets[masterWorkbook.Sheets.Count]);
                    }
                    else
                    {
                        // 如果存在，则将数据从当前工作表复制到已存在的工作表中
                        Excel.Range sourceRange = originalSh.UsedRange;
                        sourceRange = sourceRange.Offset[1, 0].Resize[sourceRange.Rows.Count - 1, sourceRange.Columns.Count];
                        Excel.Range destinationRange = existingSheet.Cells[existingSheet.UsedRange.Rows.Count + 1, 1];
                        sourceRange.Copy(destinationRange);
                    }

                    originalWb.Close(false);
                    Foundation.ReleaseObject(originalWb);
                }

                if (productLevelWb.Sheets.Count == 1)
                {
                    Excel.Worksheet sh = productLevelWb.Sheets.Add();
                    sh.Name = "0";
                }
                else
                {
                    FilterAndDeleteRowsWithBlankValue(productLevelWb.Sheets["0"], "优惠后价格");
                }

                if (skuLevelWb.Sheets.Count == 1)
                {
                    Excel.Worksheet sh = skuLevelWb.Sheets.Add();
                    sh.Name = "0";
                }
                else
                {
                    FilterAndDeleteRowsWithBlankValue(skuLevelWb.Sheets["0"], "优惠后价格");

                    Excel.Worksheet sh = skuLevelWb.Sheets["0"];
                    int iRow = sh.Range["A1000000"].End[Excel.XlDirection.xlUp].Row;
                    sh.Range[$"Q2:Q{iRow}"].FormulaR1C1 = Foundation.FillExcelWithJSONRules(_folderPath, "SKU级单品宝");

                    CreatePivotTable(skuLevelWb, "0", "Sheet3");
                }

                Globals.ThisAddIn.Application.DisplayAlerts = false;
                productLevelWb.SaveAs(_folderPath + "\\单品宝\\单品宝-商品级.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook);
                skuLevelWb.SaveAs(_folderPath + "\\单品宝\\单品宝-SKU级.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook);
                Globals.ThisAddIn.Application.DisplayAlerts = true;
            }
        }
    }
}
