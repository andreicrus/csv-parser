using csv_parser.Extensions;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace csv_parser.Core
{
    public class InteropExcel
    {
        public static void GetValues(string pathOfExcelFile, List<string> commonTemplateValues)
        {
            Excel.Application excelApp = new Excel.Application();

            List<List<string>[]> finalList = new List<List<string>[]>();

            excelApp.DisplayAlerts = false; 

            var filePath = Path.GetFullPath(pathOfExcelFile);

            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); 

            var fileName = Path.GetFileNameWithoutExtension(pathOfExcelFile);

            Excel.Worksheet sheet = workbook.Sheets[fileName];

            DataTable table = new DataTable();

            foreach (var item in commonTemplateValues)
            {
                var result = RetrieveColumnByHeader(sheet, item);
                finalList.Add(result);
            }
            workbook.Close();

            for(int i=0; i< finalList.Count; i++)
            {
            }

            ClosedXML.ExportToExcel(table,Path.GetFileName(pathOfExcelFile));

            //CreateCSVFile(table, ConfigurationManager.AppSettings["Output"]);
        }

        public static List<string>[] RetrieveColumnByHeader(Excel.Worksheet sheet, string FindWhat)
        {
            Excel.Range rngHeader = sheet.Rows[1] as Excel.Range;

            int rowCount = sheet.UsedRange.Rows.Count;
            int columnCount = sheet.UsedRange.Columns.Count;
            int index = 0;

            Excel.Range rngResult = null;
            string FirstAddress = null;

            List<string>[] columnValue = new List<string>[columnCount];

            rngResult = rngHeader.Find(What: FindWhat, LookIn: Excel.XlFindLookIn.xlValues,
            LookAt: Excel.XlLookAt.xlWhole, SearchOrder: Excel.XlSearchOrder.xlByColumns, MatchCase: true);

            if (rngResult != null)
            {
                FirstAddress = rngResult.Address;
                Excel.Range cRng = null;
                do
                {
                    columnValue[index] = new List<string>();
                    for (int i = 1; i <= rowCount; i++)
                    {
                        cRng = sheet.Cells[i, rngResult.Column] as Excel.Range;
                        if (cRng.Value != null)
                        {
                            columnValue[index].Add(cRng.Value.ToString());
                        }
                    }

                    index++;
                    rngResult = rngHeader.FindNext(rngResult);
                } while (rngResult != null && rngResult.Address != FirstAddress);

            }
            Array.Resize(ref columnValue, index);
            return columnValue;
        }

        public static void CreateDataTable(List<string>[] result,string fileName, DataTable table)
        {
            System.Data.DataColumn newColumn = new System.Data.DataColumn();
            newColumn.DefaultValue = result;
            table.Columns.Add(newColumn);
        }

        public static void CreateCSVFile(DataTable dt, string strPath)
        {
            var strFilePath = Path.GetFullPath(strPath);
            StreamWriter sw = new StreamWriter(strFilePath, true);
            int iColCount = dt.Columns.Count;

            for (int i = 0; i < iColCount; i++)
            {
                sw.Write(dt.Columns[i]);
                if (i < iColCount - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);

            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < iColCount; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        sw.Write(dr[i].ToString());
                    }
                    if (i < iColCount - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
    }
}
