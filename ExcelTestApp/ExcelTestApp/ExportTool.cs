using System;
using System.Collections.Generic;

namespace ExcelTestApp
{
    public class ExportTool
    {
        private Repository _repository;

        public ExportTool()
        {
            _repository = new Repository();
        }

        public void ExportDiseases(int diseaseCount = 20)
        {
            List<Disease> diseases = _repository.GetRangeOfDiseases(diseaseCount);
            //_repository.UpdateExportedDiseases(orphaNumbers FROM diseases retrieved ABOVE^^^^^);
            TranslationMapper translations = new TranslationMapper();

            ExcelDataManager<Disease> dataManager = new ExcelDataManager<Disease>(translations);

            foreach (Disease disease in diseases)
            {
                //// Populating dummy data for disease lists
                //disease.Synonyms = Synonym.GetDummyData();
                //disease.Summaries = Summary.GetDummyData();

                dataManager.AddNewData(disease);
            }

            dataManager.SaveFile(DateTime.Now.ToString());
        }
    }
}