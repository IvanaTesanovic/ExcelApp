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

            for (var i = 0; i < 1; i++)
                caoIvana.ExportDiseases();

            //var count = caoNenade.ImportDiseases();

            //var result = caoIvana.GetDiseaseByOprha("141209");

            //Console.WriteLine($"Finished importing {count} diseases.");
            Console.WriteLine($"Finished exporting diseases.");

            Console.ReadLine();
        }
    }
}
