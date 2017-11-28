using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportTool caoIvana = new ExportTool();
            caoIvana.ExportDiseases(4);
            Console.ReadLine();
        }
    }
}
