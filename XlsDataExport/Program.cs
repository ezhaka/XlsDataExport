using System;
using System.IO;
using Newtonsoft.Json;
using XlsDataExport.Model;

namespace XlsDataExport
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceData = File.ReadAllText("source.txt");
            DataItem deserializedData = JsonConvert.DeserializeObject<DataItem>(sourceData);

            IExcelDataWriter writer = new XlsDataWriter();
            writer.Write(deserializedData, "output.xls");

            Console.WriteLine("All done");
            Console.ReadKey();
        }
    }
}
