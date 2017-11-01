using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ExcelTestApp
{
    public class ExcelTool
    {
        private string _excelTemplatePath; //template iz kog vucemo vrednosti za property-e
        private string _excelOriginalPath; //originalni excel fajl koji ce se kopirati svaki put kad se budu export-ovale nove bolesti
        private string _excelDataPath; //fajl u koji ce biti export-ovane bolesti

        //TODO: Disease data to be exported should be sent as a parameter
        //The data is from the translation page (there should be a button for exporting?)
        //Or those two things should be separated?

        public ExcelTool()
        {
            _excelOriginalPath = "C:\\Users\\i.tesanovic\\Documents\\visual studio 2015\\Projects\\ExcelTestApp\\Data\\RetkeBolestiOriginal.xlsx";
            _excelTemplatePath = "C:\\Users\\i.tesanovic\\Documents\\visual studio 2015\\Projects\\ExcelTestApp\\Data\\BazaRetkihBolestiTemplate.xlsx";
            _excelDataPath = "C:\\Users\\i.tesanovic\\Documents\\visual studio 2015\\Projects\\ExcelTestApp\\Data\\RetkeBolesti.xlsx";

            ResetData();
            //InitializeExcelDocument(5);
        }

        //Export to Excel file from Disease object
        //Here we can even have a list of diseases
        //We should export the diseases into an excel file using the given template where each disease will belong to a different sheet
        public void ExportTo(object disease)
        {
        }

        //Import from Excel file to Disease object
        public void ImportFrom(object file)
        {
        }

        private void InitializeExcelDocument(int numberOfSheets)
        {
            for (int i = 2; i < numberOfSheets + 1; i++)
            {
                InsertEmptySheet(i);
            }

            for (int i = 1; i < numberOfSheets + 1; i++)
            {
                InsertTemplateIntoSheet(i);
            }
        }

        private void InsertEmptySheet(int sheetNumber)
        {
            using (var connection = new OleDbConnection(GetConnectionString(_excelDataPath)))
            {
                connection.Open();

                var query = $"CREATE TABLE [Sheet{sheetNumber}] (Property varchar(255), English varchar(255), Serbian varchar(255))";

                using (var command = new OleDbCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private void InsertTemplateIntoSheet(int sheetNumber)
        {
            foreach (var property in GetPropertyNames())
            {
                InsertPropertyIntoSheet(sheetNumber, property);
            }
        }

        private void InsertPropertyIntoSheet(int sheetNumber, string propertyName)
        {
            using (var connection = new OleDbConnection(GetConnectionString(_excelDataPath)))
            {
                connection.Open();
                var query = $"INSERT INTO [Sheet{sheetNumber}$] (Property) VALUES ('{propertyName}')";

                using (var command = new OleDbCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        //ovde cu prolaziti kroz listu bolesti izvucenih iz baze (npr ima ih dvadeset, nije ni bitno)
        //radicu for obican da bih imala index i onda cu ubacivati u i sheet i-tu bolest
        //private void InsertDataOnEnglishIntoSheets(int sheetNumber, object disease)
        //{
        //    using(var connection = new OleDbConnection(_connectionString))
        //    {
        //        connection.Open();

        //        var query = $"INSERT INTO [Sheet{sheetNumber}$] (English) Values(2,4,5)";
        //    }
        //}

        private List<string> GetPropertyNames()
        {
            List<string> propertyNames = new List<string>();
            var dataSet = ReadPropertyNames();

            for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                propertyNames.Add(dataSet.Tables[0].Rows[i][0].ToString());

            return propertyNames;
        }

        private DataSet ReadPropertyNames()
        {
            var dataSet = new DataSet();

            using (var connection = new OleDbConnection(GetConnectionString(_excelTemplatePath)))
            {
                connection.Open();
                var query = "SELECT Property FROM [Sheet1$]";

                using (var adapter = new OleDbDataAdapter(query, connection))
                {
                    adapter.Fill(dataSet);
                }
            }

            return dataSet;
        }

        private string GetConnectionString(string filePath) => $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR=YES;'";

        private void ResetData()
        {
            if (File.Exists(_excelDataPath))
            {
                File.Delete(_excelDataPath);
            }

            File.Copy(_excelOriginalPath, _excelDataPath);
        }
    }
}