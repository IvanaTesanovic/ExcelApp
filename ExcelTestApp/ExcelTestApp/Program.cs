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
            ExcelTool tool = new ExcelTool();
            Console.WriteLine("Inserting into excel file successfully finished. Press any button to close the dialog.");
            Console.ReadLine();
        }
    }
}
