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

        private Repository _repo;
        private int _numberOfDiseases;

        //TODO: Disease data to be exported should be sent as a parameter
        //The data is from the translation page (there should be a button for exporting?)
        //Or those two things should be separated?

        public ExcelTool()
        {
            _excelOriginalPath = "C:\\Users\\i.tesanovic\\Documents\\visual studio 2015\\Projects\\ExcelTestApp\\Data\\RetkeBolestiOriginal.xlsx";
            _excelTemplatePath = "C:\\Users\\i.tesanovic\\Documents\\visual studio 2015\\Projects\\ExcelTestApp\\Data\\BazaRetkihBolestiTemplate.xlsx";
            _excelDataPath = "C:\\Users\\i.tesanovic\\Documents\\visual studio 2015\\Projects\\ExcelTestApp\\Data\\RetkeBolesti.xlsx";
            _repo = new Repository();
            _numberOfDiseases = 5;

            ResetData();
            InitializeExcelDocument(_numberOfDiseases);
            ExportTo(_numberOfDiseases);
            
            //InitializeExcelDocument(5);
        }

        //Export to Excel file from Disease object
        //Here we can even have a list of diseases
        //We should export the diseases into an excel file using the given template where each disease will belong to a different sheet
        public void ExportTo(int numberOfDiseases)
        {
            var diseases = _repo.GetRangeOfDiseases(_numberOfDiseases);

            InsertDataIntoSheet(1, diseases[0].Name);

            //for (int i = 1; i <= diseases.Count; i++)
            //{
            //    InsertDataIntoSheet(i, diseases[i - 1].Name);
            //}
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
            using (var connection = new OleDbConnection(GetConnectionString(_excelDataPath, "Yes")))
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
            using (var connection = new OleDbConnection(GetConnectionString(_excelDataPath, "Yes")))
            {
                connection.Open();
                var query = $"INSERT INTO [Sheet{sheetNumber}$] (Property) VALUES ('{propertyName}')";

                using (var command = new OleDbCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private void InsertDataIntoSheet(int sheetNumber, string propertyValue)
        {
            using (var connection = new OleDbConnection(GetConnectionString(_excelDataPath, "No")))
            {
                connection.Open();

                //moram vrednost po vrednost od bolesti da ubacujem, ne moze ovako
                //"UPDATE ["+sheetName+"$B5:B5] SET F1=17", oledbConn
                var query = $"UPDATE [Sheet{sheetNumber}$B2:B2] SET F1='{propertyValue}'";

                using (var command = new OleDbCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

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

            using (var connection = new OleDbConnection(GetConnectionString(_excelTemplatePath, "Yes")))
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

        private string GetConnectionString(string filePath, string hdr) => $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR={hdr};'";

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