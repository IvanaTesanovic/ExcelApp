using System.Collections.Generic;
using System.Linq;

namespace ExcelTestApp
{
    public class TranslationMapper
    {
        private List<TranslationModel> _translations { get; set; }

        public TranslationMapper()
        {
            _translations = new List<TranslationModel>();
            MapToEnglish(); // Mapping to english exists only to add English to list of languages to export
            MapToSerbian(); // Mapping to serbian is the main mapping that is used for excel export
        }

        private void MapToSerbian()
        {
            TranslationModel serbian = new TranslationModel("Serbian");
            serbian.MapTranslation("Name", "Pun naziv bolesti");
            serbian.MapTranslation("ShortName", "Skraćeni naziv bolesti");
            serbian.MapTranslation("Definition", "Definicija");
            serbian.MapTranslation("Synonyms", "Sinonimi");
            serbian.MapTranslation("Category", "Kategorija");
            serbian.MapTranslation("SubCategory", "Potkategorija");
            serbian.MapTranslation("ForeignName", "Naziv bolesti na stranom jeziku");
            serbian.MapTranslation("OrphaNumber", "Orpha broj");
            serbian.MapTranslation("Prevalence", "Učestalost");
            serbian.MapTranslation("Inheritance", "Nasleđivanje");
            serbian.MapTranslation("AgeOfOnSet", "Period početka bolesti");
            serbian.MapTranslation("Icd10", "ICD-10");
            serbian.MapTranslation("Omim", "OMIM");
            serbian.MapTranslation("Umls", "UMLS");
            serbian.MapTranslation("Gard", "GARD");
            serbian.MapTranslation("MedDra", "MedDRA");
            //serbian.MapTranslation("Summaries", "Sadržaj"); // Is it needed?
            serbian.MapTranslation("Clinical description", "Tekstualni opis");
            serbian.MapTranslation("Etiology", "Etiologija");
            serbian.MapTranslation("Prognosis", "Prognoza");
            serbian.MapTranslation("Differential diagnosis", "Diferencijalna dijagnoza");
            serbian.MapTranslation("Management and treatment", "Tretman");
            serbian.MapTranslation("Diagnostic methods", "Dijagnostičke metode");
            serbian.MapTranslation("Antenatal diagnosis", "Antenatalna dijagnoza");
            serbian.MapTranslation("Epidemiology", "Epidemiologija");
            serbian.MapTranslation("Genetic counseling", "Genetsko savetovanje");
            serbian.MapTranslation("Therapy", "Terapija");
            serbian.MapTranslation("Clinical Trials", "Klinička istraživanja");
            _translations.Add(serbian);
        }

        private void MapToEnglish()
        {
            TranslationModel english = new TranslationModel("English");
            english.MapTranslation("Name", "Disease name");
            english.MapTranslation("Synonyms", "Synonyms");
            english.MapTranslation("OrphaNumber", "Orpha number");
            english.MapTranslation("Prevalence", "Prevalence");
            english.MapTranslation("Inheritance", "Inheritance");
            english.MapTranslation("AgeOfOnSet", "Age of on set");
            english.MapTranslation("Summaries", "Summary");
            english.MapTranslation("Definition", "Description");
            english.MapTranslation("Icd10", "ICD-10");
            english.MapTranslation("Omim", "OMIM");
            english.MapTranslation("Umls", "UMLS");
            english.MapTranslation("Gard", "GARD");
            english.MapTranslation("MedDra", "MedDRA");
            _translations.Add(english);
        }

        public string[] GetLanguages()
        {
            return _translations.Select(t => t.Name).ToArray();
        }

        public TranslationModel GetMapper(string language)
        {
            return _translations.FirstOrDefault(t => t.Name == language);
        }
    }
}