using csv_parser.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace csv_parser
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> templateFile = CSVParser_IO.ReadCSV(ConfigurationManager.AppSettings["TemplateFile"]);
            List<string> headerTemplate = new List<string>();
            List<string> commonHeaderTemplate = new List<string>();

            for (int i = 0; i < 69; i++)
            {
                headerTemplate.Add(templateFile[i]);
            }

            foreach (string column in headerTemplate)
            {
                commonHeaderTemplate.Add(column);
            } 

            foreach (string file in Directory.EnumerateFiles(ConfigurationManager.AppSettings["PremierLeagueData"], "*.csv"))
            {
                List<string> dataValues = CSVParser_IO.ReadCSV(file);
                
                foreach(string column in headerTemplate)
                {
                    if(!dataValues.Contains(column))
                    {
                        commonHeaderTemplate.Remove(column);
                    }
                }
            }

            //ManipulateColumns(commonHeaderTemplate);
            RemoveColumns(commonHeaderTemplate);

            Console.WriteLine("the end");
            Console.ReadLine();
        }

        internal static void ManipulateColumns(List<string> commonHeaderTemplate)
        {
            foreach (string file in Directory.EnumerateFiles(ConfigurationManager.AppSettings["PremierLeagueData"], "*.csv"))
            {
                InteropExcel.GetValues(file, commonHeaderTemplate);
            }
        }

        internal static void RemoveColumns(List<string> commonHeaderTemplate)
        {
            foreach (string file in Directory.EnumerateFiles(ConfigurationManager.AppSettings["PremierLeagueData"], "*.csv"))
            {
                List<string> fileColumns = CSVParser_IO.ReadCSV(file);
                List<string> fileHeaderColumn = new List<string>();

                int numberofcolumns = InteropExcel.GetNumberOfColumns(file);

                for (int i = 0; i < numberofcolumns; i++)
                {
                    fileHeaderColumn.Add(fileColumns[i]);
                }
                List<string> extraColumns = new List<string>();
                foreach (var item in fileHeaderColumn)
                {
                    if (!commonHeaderTemplate.Contains(item))
                    {
                        extraColumns.Add(item);
                    }
                }
                InteropExcel.RemoveColumns(file, extraColumns);
            }
        }
    }
}
