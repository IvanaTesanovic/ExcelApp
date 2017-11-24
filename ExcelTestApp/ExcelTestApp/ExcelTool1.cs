using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestApp
{
    public class ExcelTool1
    {
        private Excel.Application _xlApp;
        private Repository _repository;

        public ExcelTool1()
        {
            _repository = new Repository();
            _xlApp = new Excel.Application();

            if (_xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }
        }

        public void ExportData(int diseaseCount = 1)
        {
            object misValue = System.Reflection.Missing.Value;
            Excel.Workbook xlWorkBook = _xlApp.Workbooks.Add(misValue);

            List<Disease> diseases = _repository.GetRangeOfDiseases(diseaseCount);

            foreach (Disease disease in diseases)
            {
                // Populating dummy data for disease lists
                disease.Synonyms = Synonym.getDummyData();
                disease.Summaries = Summary.getDummyData();

                Excel.Worksheet xlWorkSheet = ExcelDocumentManager.AddNewSheet(xlWorkBook, disease.OrphaNumber);

                ExcelDocumentManager.InsertHeaders(xlWorkSheet);
                ExcelDocumentManager.PopulateFieldNames<Disease>(xlWorkSheet);
                ExcelDocumentManager.PopulateFieldData<Disease>(xlWorkSheet, disease);
                ExcelDocumentManager.ApplyStyles(xlWorkSheet);
            }

            ExcelDocumentManager.DeleteFirstSheet(xlWorkBook);

            DateTime time = DateTime.Now;

            try
            {
                xlWorkBook.SaveAs($"C:\\Users\\n.percic\\Desktop\\ExcelApp\\ExcelTestApp\\Data\\OBRB-{time}.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                _xlApp.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.FinalReleaseComObject(_xlApp);
            }

            Console.WriteLine($"Excel file created , you can find the file C:\\Users\\n.percic\\Desktop\\ExcelApp\\ExcelTestApp\\Data\\OBRB-{time}.xls");
        }

        private class ExcelDocumentManager
        {
            public static Excel.Worksheet AddNewSheet(Excel.Workbook xlWorkBook, string name = null)
            {
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);

                if(name != null)
                    xlWorkSheet.Name = name;

                return xlWorkSheet;
            }

            public static void InsertHeaders(Excel.Worksheet xlWorkSheet)
            {
                xlWorkSheet.Rows[1].Insert();
                xlWorkSheet.Cells[1, 1] = "Property";
                xlWorkSheet.Cells[1, 2] = "English";
                xlWorkSheet.Cells[1, 3] = "Serbian";
            }

            public static void PopulateFieldNames<T>(Excel.Worksheet xlWorkSheet, int startingRow = 2) where T: new()
            {
                var props = typeof(T).GetProperties().ToArray();
                for (var i = 0; i < props.Length; i++)
                {
                    xlWorkSheet.Cells[i + startingRow, 1].Font.Bold = true;
                    xlWorkSheet.Cells[i + startingRow, 1] = props[i].Name;
                }
            }

            public static void PopulateFieldData<T>(Excel.Worksheet xlWorkSheet, T data, int startingRow = 2) where T : new()
            {
                var props = typeof(T).GetProperties().ToArray();
                for (var i = 0; i < props.Length; i++)
                {
                    if (props[i].PropertyType.IsGenericType && props[i].PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                        switch (props[i].Name)
                        {
                            case "Synonyms":
                                List<Synonym> synonyms = ((List<Synonym>)props[i].GetValue(data));
                                xlWorkSheet.Cells[i + startingRow, 2] = string.Join(Environment.NewLine, synonyms.Select(s => s.Name));
                                break;

                            case "Summaries":
                                List<Summary> summaries = ((List<Summary>)props[i].GetValue(data));
                                foreach(Summary sum in summaries)
                                {
                                    xlWorkSheet.Rows[i + startingRow].Insert();
                                    xlWorkSheet.Cells[i + startingRow, 1] = sum.Title;
                                    xlWorkSheet.Cells[i + startingRow, 2] = sum.Text;
                                }
                                break;
                        }
                    else
                        xlWorkSheet.Cells[i + startingRow, 2] = props[i].GetValue(data).ToString();
                }
            }

            public static void ApplyStyles(Excel.Worksheet xlWorkSheet)
            {
                xlWorkSheet.Rows[1].Font.Bold = true;

                // Shows compatibility error when saving excel on file system, check better solution for this styling if needed :)
                //xlWorkSheet.Rows[1].Interior.Color = Excel.XlRgbColor.rgbPaleVioletRed;
                //xlWorkSheet.Columns[1].Interior.Color = Excel.XlRgbColor.rgbPaleVioletRed;

                xlWorkSheet.Columns[1].ColumnWidth = 18;
                xlWorkSheet.Columns[2].ColumnWidth = 40;
                xlWorkSheet.Columns[3].ColumnWidth = 40;
                xlWorkSheet.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                xlWorkSheet.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                xlWorkSheet.Cells.WrapText = true;
            }

            internal static void DeleteFirstSheet(Excel.Workbook xlWorkBook)
            {
                xlWorkBook.Sheets[1].Delete();
            }
        }
    }
}