using System;
using System.Collections;

namespace ExcelTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportTool caoIvana = new ExportTool();

            //ImportTool caoNenade = new ImportTool();

            for (var i = 0; i < 49; i++)
                caoIvana.ExportDiseases();

            //var count = caoNenade.ImportDiseases();

            //var result = caoIvana.GetDiseaseByOprha("141209");

            //Console.WriteLine($"Finished importing {count} diseases.");
            Console.ReadLine();
        }
    }
}
