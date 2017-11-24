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
            ExcelTool1 caoIvana = new ExcelTool1();
            caoIvana.ExportData(6);
            Console.ReadLine();
        }
    }
}
