using ExcelTestApp.Entities;
using ExcelTestApp.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp.Business
{
    public class MapperService
    {
        private readonly Repository _repository;
        private readonly Mapper _mapper;

        public MapperService()
        {
            _repository = new Repository();
            _mapper = new Mapper();
        }

        public void InsertDisease(Dictionary<int, string> data)
        {
            InsertDisease(_mapper.MapToDisease(data));
        }

        private void InsertDisease(Disease disease)
        {
            DiseaseEntity original;

            if (disease.OrphaNumber == "")
                original = _repository.GetOriginalDiseaseByOrpha(disease.OrphaNumber);

            original = _repository.GetOriginalDiseaseByOrpha(disease.OrphaNumber);

            var entity = _mapper.MapToDiseaseEntity(disease, original.Id, 3, original.MeSH);
            var synonyms = disease.Synonyms.Select(s => _mapper.MapToSynonymEntity(s, entity.Id));
            var summaries = disease.Summaries.Select(s => _mapper.MapToSummaryEntity(s, entity.Id));

            _repository.InsertDisease(entity, synonyms, summaries);
        }
    }
}
