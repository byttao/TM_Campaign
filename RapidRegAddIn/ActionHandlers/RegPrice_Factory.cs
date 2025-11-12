using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using RapidRegAddIn.Utilities;

namespace RapidRegAddIn.ActionHandlers
{
    public static class RegPrice_Factory
    {
        private static Workbook _wb;
        private static Worksheet _sh;

        public static void CreateParams(string path)
        {
            _wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            _sh = _wb.ActiveSheet;

            Globals.ThisAddIn.Application.ScreenUpdating = false;
            try
            {
                PriceCalc();
            }
            catch (Exception e)
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                MessageBox.Show(e.ToString());
                throw;
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private static void PriceCalc()
        {
            // 复制第二行
            Range sourceRange = _sh.Rows[2];
            sourceRange.Copy();

            // 将复制的行插入到第四行
            Range destinationRange = _sh.Rows[4];
            destinationRange.Insert(XlInsertShiftDirection.xlShiftDown);

            // 清除筛选
            _sh.AutoFilterMode = false;
            //开启筛选
            _sh.Rows["4:4"].AutoFilter();
            int iRow = _sh.UsedRange.Rows.Count;
            Range rng = _sh.Range[$"A5:D{iRow}"];
            //取消合并单元格
            rng.UnMerge();
            //定位空值并填充公式
            Range specialRange = rng.SpecialCells(XlCellType.xlCellTypeBlanks);

            if (specialRange != null)
            {
                specialRange.FormulaR1C1 = "=R[-1]C";
                // 设置单元格格式为文本
                rng.NumberFormat = "@";

                rng.Value2 = rng.Value2;
            }

            //一口价数值化
            _sh.Range["G1"].UnMerge();
            _sh.Range["G:G"].TextToColumns(Comma: false, ConsecutiveDelimiter: false, DataType: XlTextParsingType.xlDelimited, Destination: _sh.Range["G:G"], Other: false,
                Semicolon: false, Space: false, Tab: false, TextQualifier: XlTextQualifier.xlTextQualifierNone);

            //选择数据透视区域
            rng = _sh.Range[$"A4:G{iRow}"];

            // 添加一个新的工作表用于放置透视表
            Worksheet pivotSheet = _wb.Worksheets.Add();
            pivotSheet.Name = "商品级";
            // 定义透视表的目标位置
            Range pivotDestination = pivotSheet.Cells[2, 1];

            // 创建 PivotTable
            PivotTable pivotTable = pivotSheet.PivotTableWizard(XlPivotTableSourceType.xlDatabase,
                rng,
                pivotDestination,
                "PivotTable");

            // 设置透视表字段

            pivotTable.PivotFields("商品ID").Orientation = XlPivotFieldOrientation.xlRowField;
            pivotTable.AddDataField(pivotTable.PivotFields("一口价"), "最小值项:一口价", XlConsolidationFunction.xlMin);
            pivotTable.AddDataField(pivotTable.PivotFields("一口价"), "最大值项:一口价", XlConsolidationFunction.xlMax);
            pivotTable.DataPivotField.Orientation = XlPivotFieldOrientation.xlColumnField;
            pivotTable.PivotFields("商品ID").Position = 1;
            pivotTable.DataFields["最小值项:一口价"].NumberFormat = "#,##0.00";
            pivotTable.DataFields["最大值项:一口价"].NumberFormat = "#,##0.00";

            pivotSheet.Range["D3:L3"].Value2 = new[]
            {
                "一口价",
                "商品级价格",
                "SKU最低",
                "SKU最高",
                "差异",
                "分类",
                "活动基准价",
                "商品ID",
                "活动报名价"
            };

            pivotSheet.Range["D4"].FormulaR1C1 = "=IF(RC[-2]=RC[-1],RC[-2],RC[-2]&\"~\"&RC[-1])";
            pivotSheet.Range["E4"].FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],'[单品宝-商品级.xlsx]0'!C10:C15,6,0),\"\")";
            pivotSheet.Range["F4"].FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],'[单品宝-SKU级.xlsx]Sheet3'!C1:C3,2,0),\"\")";
            pivotSheet.Range["G4"].FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],'[单品宝-SKU级.xlsx]Sheet3'!C1:C3,3,0),\"\")";
            pivotSheet.Range["H4"].FormulaR1C1 = "=IF(RC[-2]<>\"\",RC[-2]=RC[-1],\"\")";
            pivotSheet.Range["I4"].FormulaR1C1 =
                "=IF(OR(RC[-4]<>\"\",RC[-1]),IF(RC[-4]=RC[1],\"商品级-有单品宝\",\"SKU级-有单品宝\"),IF(RC[-1]<>\"\",\"SKU级-有单品宝\",IF(ISERROR(FIND(\"~\",RC[-5])),\"商品级-【无单品宝】\",\"SKU级-【无单品宝】\")))";
            pivotSheet.Range["J4"].FormulaR1C1 = "=MIN(RC[-6],RC[-5],RC[-4])";
            pivotSheet.Range["K4"].FormulaR1C1 = "=RC[-10]";
            pivotSheet.Range["L4"].FormulaR1C1 = Foundation.FillExcelWithJSONRules(Path.GetDirectoryName(_wb.FullName), "草稿-商品级");

            int pRow = pivotSheet.Range["A1000000"].End[XlDirection.xlUp].Row;
            rng = pivotSheet.Range[$"D4:L{pRow}"];
            pivotSheet.Range["D4:L4"].AutoFill(rng, XlAutoFillType.xlFillSeries);

            _sh.Columns["I:J"].Insert(XlInsertShiftDirection.xlShiftToRight);
            rng = _sh.Range["I4:J4"];
            rng.Value2 = new[]
            {
                "类别",
                "活动价"
            };
            rng.Font.Bold = true;
            rng.Interior.Color = XlRgbColor.rgbYellow;
            rng.Font.Color = XlRgbColor.rgbRed;
            _sh.Range["I5"].FormulaR1C1 = "=VLOOKUP(RC[-8],商品级!C1:C10,9,0)";
            _sh.Range["J5"].FormulaR1C1 = Foundation.FillExcelWithJSONRules(Path.GetDirectoryName(_wb.FullName), "草稿-活动价");

            rng = _sh.Range[$"I5:J{iRow}"];
            _sh.Range["I5:J5"].AutoFill(rng, XlAutoFillType.xlFillSeries);

            //差额
        }
    }
}
