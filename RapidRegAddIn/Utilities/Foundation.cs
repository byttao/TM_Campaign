using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RapidRegAddIn.Utilities
{
    public static class Foundation
    {
        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        public static string FillExcelWithJSONRules(string jsonPath,string fieldname)
        {
            
            // 读取 JSON 文件内容
            string jsonContent = File.ReadAllText(jsonPath+"\\relus.json");

            // 将 JSON 内容转换为对象
            JObject jsonData = JObject.Parse(jsonContent);
            string r1c1= jsonData["rules"].First(c=>c["fieldname"].ToString()==fieldname)["formulaR1C1"].ToString();
            return r1c1;
        }
    }
}
