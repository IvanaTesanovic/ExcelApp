﻿using ExcelTestApp.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelTestApp.Helper
{
    public class Mapper
    {
        public DiseaseEntity MapToDiseaseEntity(Disease disease, Guid parentDisease, int state, string mesh = null)
        {
            return new DiseaseEntity
            {
                Id = Guid.NewGuid(),
                AgeOfOnset = GetTrimmedValue(disease.AgeOfOnset),
                Definition = GetTrimmedValue(disease.Definition),
                Gard = GetTrimmedValue(disease.Gard),
                Icd10 = GetTrimmedValue(disease.Icd10),
                Inheritance = GetTrimmedValue(disease.Inheritance),
                MedDra = GetTrimmedValue(disease.MedDra),
                Name = GetTrimmedValue(disease.Name),
                Omim = GetTrimmedValue(disease.Omim),
                OrphaNumber = GetTrimmedValue(disease.OrphaNumber),
                Prevalence = GetTrimmedValue(disease.Prevalence),
                Umls = GetTrimmedValue(disease.Umls),
                MeSH = GetTrimmedValue(mesh) ?? GetTrimmedValue(disease.MeSH),
                IsTranslationOf = parentDisease,
                State = state
            };
        }

        public SynonymEntity MapToSynonymEntity(Synonym synonym, Guid diseaseId)
        {
            return new SynonymEntity
            {
                DiseaseId = diseaseId,
                Id = Guid.NewGuid(),
                Name = GetTrimmedValue(synonym.Name)
            };
        }

        public SummaryEntity MapToSummaryEntity(Summary summary, Guid diseaseId)
        {
            return new SummaryEntity
            {
                Id = Guid.NewGuid(),
                DiseaseId = diseaseId,
                Text = GetTrimmedValue(summary.Text),
                Title = GetTrimmedValue(summary.Title)
            };
        }

        public List<Disease> MapToDiseases(List<Dictionary<int, string>> data)
        {
            var diseases = new List<Disease>();

            foreach (var row in data)
            {
                diseases.Add(MapToDisease(row));
            }

            return diseases;
        }

        public Disease MapToDisease(Dictionary<int, string> data)
        {
            var disease = new Disease
            {
                AgeOfOnset = data[12],
                Definition = data[4],
                Gard = data[17],
                Icd10 = data[13],
                Inheritance = data[11],
                MedDra = data[18],
                MeSH = data[16], //this doesn't exist in the sheet.
                Name = data[2],
                Omim = data[14],
                OrphaNumber = data[9],
                Prevalence = data[10],
                Umls = data[15],
                Synonyms = MapSynonyms(data[5]),
                Summaries = MapSummaries(data)
            };

            return disease;
        }

        private IEnumerable<Summary> MapSummaries(Dictionary<int, string> data)
        {
            var summaries = new List<Summary>();

            summaries.Add(new Summary { Title = "Tekstualni opis", Text = data[19] });
            summaries.Add(new Summary { Title = "Etiologija", Text = data[20] });
            summaries.Add(new Summary { Title = "Prognoza", Text = data[21] });
            summaries.Add(new Summary { Title = "Diferencijalna dijagnoza", Text = data[22] });
            summaries.Add(new Summary { Title = "Tretman", Text = data[23] });
            summaries.Add(new Summary { Title = "Dijagnostičke metode", Text = data[24] });
            summaries.Add(new Summary { Title = "Antenatalna dijagnoza", Text = data[25] });
            summaries.Add(new Summary { Title = "Epidemiologija", Text = data[26] });
            summaries.Add(new Summary { Title = "Genetsko savetovanje", Text = data[27] });
            summaries.Add(new Summary { Title = "Terapija", Text = data[28] });
            summaries.Add(new Summary { Title = "Klinička istraživanja", Text = data[29] });

            return summaries.Where(s => !string.IsNullOrEmpty(s.Text));
        }

        private List<Synonym> MapSynonyms(string synonyms)
        {
            return SplitByNewLine(synonyms).Where(s => !string.IsNullOrEmpty(s)).Select(s => new Synonym(s)).ToList();
        }

        /*
         using (System.IO.StringReader reader = new System.IO.StringReader(input)) {
    string line = reader.ReadLine();
}
             */

        private string[] SplitByNewLine(string value)
        {
            return value.Split(
                    new[] { "\r\n", "\r", "\n" },
                    StringSplitOptions.None);
        }

        private string GetTrimmedValue(string value)
        {
            return Regex.Replace(value.TrimStart().TrimEnd(), @"\s+", " ");
        }
    }
}