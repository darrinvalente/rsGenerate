using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Data;

namespace RSGenerate
{
    public static class ExcelHelper
    {
        public static string GetCellValue(DataRow row, string col, bool sanitize = true)
        {
            var value = row[col].ToString();

            if (sanitize)
                return value.Replace('-', '_');
            else
                return value;
        }
    }
}
