using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp
{
    public class Disease
    {
        public string Name { get; set; }

        public string Definition { get; set; }

        public string OrphaNumber { get; set; }

        public string Prevalence { get; set; }

        public string Inheritance { get; set; }

        public string AgeOfOnset { get; set; }

        public string Icd10 { get; set; }

        public string Omim { get; set; }

        public string Umls { get; set; }

        public string MeSH { get; set; }

        public string Gard { get; set; }

        public string MedDra { get; set; }

        public List<Synonym> Synonyms { get; set; }

        public List<Summary> Summaries { get; set; }

        public override string ToString()
        {
            return OrphaNumber;
        }
    }
}
