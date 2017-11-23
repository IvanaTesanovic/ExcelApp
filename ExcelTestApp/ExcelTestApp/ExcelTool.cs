using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace ExcelTestApp
{
    public class ExcelTool
    {
        private string _excelFolderPath;

        private string _excelTemplatePath; //template iz kog vucemo vrednosti za property-e
        private string _excelOriginalPath; //originalni excel fajl koji ce se kopirati svaki put kad se budu export-ovale nove bolesti
        private string _excelDataPath; //fajl u koji ce biti export-ovane bolesti

        private Repository _repo;
        private int _numberOfDiseases;

        private List<RowPropertyEntity> RowPropertyValues = new List<RowPropertyEntity>();
        private List<Disease> _diseases = new List<Disease>();

        //TODO: _numberOfDiseases should be sent from the GUI.

        public ExcelTool()
        {
            InitializeList();

            _excelFolderPath = "C:\\Users\\n.percic\\Desktop\\ExcelApp\\ExcelTestApp\\Data\\";
            _excelOriginalPath = $"{_excelFolderPath}RetkeBolestiOriginal.xlsx";
            _excelTemplatePath = $"{_excelFolderPath}BazaRetkihBolestiTemplate.xlsx";
            _excelDataPath = $"{_excelFolderPath}RetkeBolesti.xlsx";

            _repo = new Repository();
            _numberOfDiseases = 5;

            _diseases = _repo.GetRangeOfDiseases(_numberOfDiseases);

            ResetData();
            Console.WriteLine("Data is reset.");
            InitializeExcelDocument(_numberOfDiseases);
            Console.WriteLine("Document is initialized.");
            //ExportTo(_numberOfDiseases);
            //Console.WriteLine("Data is exported.");
        }

        //kad dobavim bolesti, mozda odmah da im izracunam i karaktere i sve
        //napravim neku novu klasu koja ce biti samo za popunjavanje excel-a
        //i prolazim kroz tu novokreiranu listu i tako popunjavam excel
        //da ne moram da radim dupli posao
        //tu listu moram pre svega da popunim da bih znala koliko mi treba redova od svega
        //znaci od sinonima, sadrzaja i opisa bolesti

        //Export to Excel file from Disease object
        //Here we can even have a list of diseases
        //We should export the diseases into an excel file using the given template where each disease will belong to a different sheet
        public void ExportTo(int numberOfDiseases)
        {
            for (int i = 1; i <= _diseases.Count; i++)
            {
                var disease = _diseases[i - 1];

                var longPropertyNames = new List<string>();
                var longPropertyValues = new List<object>();

                foreach (var item in RowPropertyValues)
                {
                    var value = disease.GetType().GetProperty(item.PropertyName).GetValue(disease).ToString();

                    if (value.Length <= 255 && item.PropertyName != "Synonyms" && item.PropertyName != "Summaries")
                        InsertDataIntoSheet(i, value, item.Row);
                    else
                    {
                        longPropertyNames.Add(item.PropertyName);
                        //a sta ako je lista? onda moram prolaziti kroz listu i kreirati to, u pm
                        longPropertyValues.Add(value);
                    }
                }

                WritePropertiesToTxtFile(disease.OrphaNumber, longPropertyNames, longPropertyValues);
            }
        }

        public void ImportFrom(object file)
        {
            //TODO: Import from Excel file to Disease object
        }

        private void InitializeExcelDocument(int numberOfSheets)
        {
            for (int i = 2; i < numberOfSheets + 1; i++)
            {
                InsertEmptySheet(i);
            }

            //for (int i = 1; i < numberOfSheets + 1; i++)
            //{
            //    InsertTemplateIntoSheet(i);
            //}
        }

        private void InsertEmptySheet(int sheetNumber)
        {
            using (var connection = new OleDbConnection(GetConnectionString(_excelDataPath, "Yes")))
            {
                connection.Open();

                var query = $"CREATE TABLE [Sheet{sheetNumber}])";

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

        private void InsertDataIntoSheet(int sheetNumber, string propertyValue, int rowNumber)
        {
            using (var connection = new OleDbConnection(GetConnectionString(_excelDataPath, "No")))
            {
                connection.Open();
                var query = $"UPDATE [Sheet{sheetNumber}$B{rowNumber}:B{rowNumber}] SET F1='{propertyValue}'";

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

        //private List<ExcelDiseaseEntity> InitializeDiseases()
        //{
        //    List<ExcelDiseaseEntity> excelDiseases = new List<ExcelDiseaseEntity>();

        //    foreach(var disease in _diseases)
        //    {
        //        foreach(var property in disease.GetType().GetProperties())
        //        {
        //            var rowIndex = RowPropertyValues.Where(v => v.PropertyName == property.Name).FirstOrDefault() == null? 0 : RowPropertyValues.Where(v => v.PropertyName == property.Name).First().Row;

        //            if (property.Name != "Summaries" && property.Name != "Synonyms")
        //            {
        //                var propertyValue = property.GetValue(disease).ToString();

        //                if (propertyValue.Length > 255)
        //                {
        //                    //ovde treba duzinu podeliti sa 255 da bih znala koliko da ubacim redova
        //                    var numberOfRows = property.GetValue(disease).ToString().Length % 255 != 0 ? property.GetValue(disease).ToString().Length % 255 + 1 : property.GetValue(disease).ToString().Length % 255;
        //                    for(int i = 0; i < numberOfRows; i++)
        //                        excelDiseases.Add(new ExcelDiseaseEntity { NumberOfRows = numberOfRows, PropertyName = property.Name, RowIndex = rowIndex });
        //                }
        //                else
        //                {
        //                    //ako nije veca od 255 samo ubaciti jedan red
        //                    //rowindex cemo izvuci iz ove liste, mozda bolje da napravim dictionary gde je property name key i onda samo uzmem row index kao vrednost
        //                    //na sta ovo lici, ivana, pobogu
        //                    excelDiseases.Add(new ExcelDiseaseEntity { NumberOfRows = 1, PropertyName = property.Name, PropertyValue = propertyValue, RowIndex = rowIndex });
        //                }
        //            }
        //        }
        //    }

        //    return excelDiseases;
        //}

        //trebace se posle svaki rowindex update-ovati u zavisnosti od naziva property-a

        private void InitializeList()
        {
            RowPropertyValues.Add(new RowPropertyEntity { Row = 2, PropertyName = "Name" });
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 3, PropertyName = "" }); //skracen naziv bolesti
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 4, PropertyName = "Synonyms" }); //sinonimi
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 5, PropertyName = "" }); //kategorija
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 6, PropertyName = "" }); //podkategorija
            RowPropertyValues.Add(new RowPropertyEntity { Row = 7, PropertyName = "Name" }); //naziv na engleskom jeziku
            RowPropertyValues.Add(new RowPropertyEntity { Row = 8, PropertyName = "OrphaNumber" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 9, PropertyName = "Prevalence" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 10, PropertyName = "Inheritance" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 11, PropertyName = "AgeOfOnset" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 12, PropertyName = "Icd10" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 13, PropertyName = "Omim" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 14, PropertyName = "Umls" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 15, PropertyName = "MeSH" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 16, PropertyName = "Gard" });
            RowPropertyValues.Add(new RowPropertyEntity { Row = 17, PropertyName = "MedDra" });
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 18, PropertyName = "Summaries" }); //sadrzaj
            RowPropertyValues.Add(new RowPropertyEntity { Row = 19, PropertyName = "Definition" });
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 20, PropertyName = "" }); //dijagnostika
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 21, PropertyName = "" }); //terapije
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 22, PropertyName = "" }); //validacije lekara
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 23, PropertyName = "" }); //prognoze
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 24, PropertyName = "" }); //klinicka istrazivanja
            //RowPropertyValues.Add(new RowPropertyEntity { Row = 25, PropertyName = "" }); //dodatni clanci, linkovi, komentari
        }

        private void WritePropertiesToTxtFile(string orphaNumber, List<string> propertyNames, List<object> propertyValues)
        {
            string content = string.Empty;

            for (int i = 0; i < propertyNames.Count; i++)
            {
                content += propertyNames[i] + Environment.NewLine + propertyValues[i] + Environment.NewLine;
            }

            File.WriteAllText($"{_excelFolderPath}{orphaNumber}.txt", content);
        }

        private string GetConnectionString(string filePath, string hdr) => $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR={hdr};'";

        private void ResetData()
        {
            DeleteTxtFiles();

            if (File.Exists(_excelDataPath))
            {
                File.Delete(_excelDataPath);
            }

            File.Copy(_excelOriginalPath, _excelDataPath);
        }

        private void DeleteTxtFiles()
        {
            foreach (string file in Directory.GetFiles(_excelFolderPath, "*.txt").Where(item => item.EndsWith(".txt")))
            {
                File.Delete(file);
            }
        }

        private class RowPropertyEntity
        {
            public int Row { get; set; }
            public string PropertyName { get; set; }
        }
    }
}