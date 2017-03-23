using csv_parser.Core;
using System;

namespace csv_parser
{
    class Program
    {
        static void Main(string[] args)
        {
            var result = CSVParser_IO.ReadCSV(@".\Data\05-06.csv");
            Console.ReadLine();
        }
    }
}
