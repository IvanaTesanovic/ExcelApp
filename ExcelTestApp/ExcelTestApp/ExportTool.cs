using ExcelTestApp.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTestApp
{
    public class ExportTool
    {
        private Repository _repository;
        private List<DiseaseEntity> _exportedDiseases;

        public ExportTool()
        {
            _repository = new Repository();
            _exportedDiseases = new List<DiseaseEntity>();
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

            dataManager.SaveFile(DateTime.Now.ToString("yyyyMMddHHmmss"));
            //_repository.Update(_exportedDiseases);

        }

        public List<DiseaseEntity> GetDiseaseByOprha(string orpha)
        {
            return _repository.GetDiseaseByOrpha(orpha);
        }

        private List<Disease> GetDiseasesForExport(int diseaseCount)
        {
            List<DiseaseEntity> result = _repository.GetDiseases(diseaseCount);

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
            disease.State = 1;
            _exportedDiseases.Add(disease);

            return result;
        }
    }
}