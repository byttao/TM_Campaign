using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace RapidRegAddIn.ActionHandlers
{
    public static class QuickAction
    {
        
        private static Workbook _wb;
        private static Worksheet _sh;
        public static void CreateParams(string path)
        {
            _wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            _sh = _wb.ActiveSheet;
            jichucaozuo();
        }
        
        private static void jichucaozuo()
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

        }
    }
}
