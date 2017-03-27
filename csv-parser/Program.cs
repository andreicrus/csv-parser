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

            foreach (string file in Directory.EnumerateFiles(ConfigurationManager.AppSettings["PremierLeagueData"], "*.csv"))
            {
                InteropExcel.GetValues(file, commonHeaderTemplate);
            }

            Console.WriteLine("the end");
            Console.ReadLine();
        }
    }
}
