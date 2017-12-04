using System;
using System.Collections;

namespace ExcelTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportTool caoIvana = new ExportTool();

            for (var i = 0; i < 50; i++)
                caoIvana.ExportDiseases();

            //var result = caoIvana.GetDiseaseByOprha("141209");

            Console.WriteLine("Finished");
            Console.ReadLine();
        }
    }
}
