using System;
using System.Collections;

namespace ExcelTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportTool caoIvana = new ExportTool();

            //for(var i = 0; i < 50; i++)
                caoIvana.ExportDiseases();

            Console.ReadLine();
        }
    }
}
