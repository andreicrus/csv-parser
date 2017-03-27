using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csv_parser.Core
{
    public class ClosedXML
    {
        public static void ExportToExcel(System.Data.DataTable dataTable, string fileName)
        {
            var wb = new XLWorkbook();

            // Add a DataTable as a worksheet
            wb.Worksheets.Add(dataTable);

            wb.SaveAs(fileName);
        }
    }
}
