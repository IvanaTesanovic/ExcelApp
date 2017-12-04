using ExcelTestApp.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

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
            List<Disease> diseases = GetDiseasesForExport(diseaseCount);

            TranslationMapper translations = new TranslationMapper();
            ExcelDataManager<Disease> dataManager = new ExcelDataManager<Disease>(translations);

            foreach(Disease disease in diseases)
            {
                dataManager.AddNewData(disease);
            }

            dataManager.SaveFile(DateTime.Now.ToString());
        }

        public List<DiseaseEntity> GetDiseaseByOprha(string orpha)
        {
            return _repository.GetDiseaseByOrpha(orpha);
        }

        private List<Disease> GetDiseasesForExport(int diseaseCount)
        {
            List<DiseaseEntity> result = _repository.GetDiseases(diseaseCount);
            string[] diseaseIds = result.Select(d => d.Id.ToString()).ToArray();
            _repository.UpdateExportedDiseases(diseaseIds);
            return result.Select(d => PopulateDiseaseData(d)).ToList();
        }

        private Disease PopulateDiseaseData(DiseaseEntity disease)
        {
            Disease result = new Disease(disease);

            result.Synonyms = _repository.GetSynonymsByDiseaseId(disease.Id.ToString())
                                         .Select(s => new Synonym(s))
                                         .ToList();

            result.Summaries = _repository.GetSummariesByDiseaseId(disease.Id.ToString())
                                          .Select(s => new Summary(s))
                                          .ToList();
            return result;
        }
    }
}