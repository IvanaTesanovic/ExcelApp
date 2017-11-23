using System;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestApp
{
    public class ExcelTool1
    {
        private Excel.Application _xlApp;

        public ExcelTool1()
        {
            _xlApp = new Excel.Application();

            if (_xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }
        }

        public void ExportData(int diseaseCount = 1)
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = _xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);

            var props = typeof(Disease).GetProperties().ToArray();

            for(var i = 1; i < props.Length; i++)
            {
                xlWorkSheet.Columns[i].ColumnWidth = 18;
                xlWorkSheet.Cells[1, i] = props[i].Name;
            }


            try
            {
                xlWorkBook.SaveAs("C:\\Users\\n.percic\\Desktop\\ExcelApp\\ExcelTestApp\\Data\\test-data.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                _xlApp.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(_xlApp);
            }

            Console.WriteLine("Excel file created , you can find the file C:\\Users\\n.percic\\Desktop\\ExcelApp\\ExcelTestApp\\Data\\test-data.xls");
        }
    }
}