using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestApp
{
    public class ExportTool
    {
        private Repository _repository;

        public ExportTool()
        {
            _repository = new Repository();
        }

        public void ExportDiseases(int diseaseCount = 1)
        {
            List<Disease> diseases = _repository.GetRangeOfDiseases(diseaseCount);
            TranslationMapper translations = new TranslationMapper();

            ExcelDataManager<Disease> dataManager = new ExcelDataManager<Disease>(translations);

            foreach (Disease disease in diseases)
            {
                // Populating dummy data for disease lists
                disease.Synonyms = Synonym.GetDummyData();
                disease.Summaries = Summary.GetDummyData();

                dataManager.AddNewData(disease);
            }

            dataManager.SaveFile(DateTime.Now.ToString());
        }
    }
}