using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestApp
{
    public class ExcelDataManager<T> where T : class
    {
        private Excel.Application _xlApp;
        private Excel.Workbook _xlWorkBook;
        private TranslationMapper _translations;
        private string _primaryLanguage;

        public ExcelDataManager()
        {
        }

        public ExcelDataManager(TranslationMapper translations, string primaryLanguage = "Serbian")
        {
            _xlApp = new Excel.Application();
            if (_xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!");
                return;
            }
            _xlWorkBook = _xlApp.Workbooks.Add(Type.Missing);

            _translations = translations;
            _primaryLanguage = primaryLanguage;
        }

        private Excel.Worksheet AddNewSheet()
        {
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Sheets.Add(After: _xlWorkBook.Sheets[_xlWorkBook.Sheets.Count]);
            xlWorkSheet.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            xlWorkSheet.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            xlWorkSheet.Cells.WrapText = true;

            return xlWorkSheet;
        }

        private void InsertHeaders(Excel.Worksheet xlWorkSheet)
        {
            string[] languages = _translations.GetLanguages();

            xlWorkSheet.Rows[1].Insert();
            xlWorkSheet.Rows[1].Font.Bold = true;

            xlWorkSheet.Cells[1, 1] = "Fields";

            for (int i = 0; i < languages.Length; i++)
            {
                xlWorkSheet.Cells[1, 2 + i] = languages[i];
                xlWorkSheet.Columns[2 + i].ColumnWidth = 40;
            }

            var props = _translations.GetMapper(_primaryLanguage).GetWords();

            for (var i = 0; i < props.Length; i++)
            {
                xlWorkSheet.Cells[2 + i, 1].Font.Bold = true;
                xlWorkSheet.Cells[2 + i, 1] = props[i];
            }

            xlWorkSheet.Columns[1].ColumnWidth = 18;
        }

        private void PopulateFieldData(Excel.Worksheet xlWorkSheet, T data)
        {
            var language = _translations.GetMapper(_primaryLanguage);
            var entities = language.GetMappedEntities();
            var dataProperties = typeof(T).GetProperties();

            for (var i = 0; i < entities.Length; i++)
            {
                string cellInfo = null;
                foreach (var property in dataProperties)
                {
                    if (property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                    {
                        if (property.Name == "Synonyms")
                        {
                            if (property.Name == entities[i])
                            {
                                List<Synonym> synonyms = property.GetValue(data) as List<Synonym>;
                                cellInfo = string.Join(Environment.NewLine, synonyms.Select(s => s.Name));
                            }
                        }

                        if (property.Name == "Summaries")
                        {
                            List<Summary> summaries = property.GetValue(data) as List<Summary>;
                            foreach (var summary in summaries)
                            {
                                if (summary.Title == entities[i])
                                    cellInfo = summary.Text;
                            }
                        }
                    }
                    else
                    {
                        if (property.Name == entities[i])
                        {
                            cellInfo = property.GetValue(data).ToString();
                        }
                    }
                }

                xlWorkSheet.Cells[2 + i, 2] = cellInfo ?? "-";
            }
        }

        private void ApplyStyles(Excel.Worksheet xlWorkSheet)
        {
            // Shows compatibility error when saving excel on file system, check better solution for this styling if needed :)
            xlWorkSheet.Rows[1].Interior.Color = Excel.XlRgbColor.rgbPaleVioletRed;
            xlWorkSheet.Columns[1].Interior.Color = Excel.XlRgbColor.rgbPaleVioletRed;
        }

        public void AddNewData(T data)
        {
            Excel.Worksheet xlWorkSheet = AddNewSheet();
            InsertHeaders(xlWorkSheet);
            PopulateFieldData(xlWorkSheet, data);
            //ApplyStyles(xlWorkSheet);
        }

        public void SaveFile(string name)
        {
            // Removing automatically created first sheet
            _xlWorkBook.Worksheets[1].Delete();

            try
            {
                _xlWorkBook.SaveAs($"C:\\Users\\i.tesanovic\\Desktop\\ExportedDiseases\\OBRB-{name}.xls", Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _xlWorkBook.Close(true, Type.Missing, Type.Missing);
                _xlApp.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(_xlWorkBook);
                Marshal.FinalReleaseComObject(_xlApp);
            }

            Console.WriteLine($"Excel file created, you can find the file C:\\Users\\n.percic\\Desktop\\ExcelApp\\ExcelTestApp\\Data\\OBRB-{name}.xls");
        }

        public List<Dictionary<int, string>> LoadDataFromFile(string filePath)
        {
            List<Dictionary<int, string>> data = new List<Dictionary<int, string>>();

            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                for (int i = 1; i <= 20; i++)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[i];
                    data.Add(LoadDataFromWorksheet(worksheet));
                }
            }

            return data;
        }

        private Dictionary<int, string> LoadDataFromWorksheet(ExcelWorksheet worksheet)
        {
            Dictionary<int, string> data = new Dictionary<int, string>();

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            for (int row = 2; row <= rowCount; row++)
            {
                data.Add(row, GetCellValue(worksheet, row, 3));
            }

            return data;
        }

        private string GetCellValue(ExcelWorksheet worksheet, int row, int column)
        {
            if (worksheet.Cells[row, column].Value == null)
                return string.Empty;

            return worksheet.Cells[row, column].Value.ToString();
        }
    }
}