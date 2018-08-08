using ExcelTestApp.Entities;
using System.Collections.Generic;

namespace ExcelTestApp
{
    public class Synonym
    {
        public string Name { get; set; }

        public Synonym(string name)
        {
            Name = name;
        }

        public Synonym(SynonymEntity synonym)
        {
            Name = synonym.Name;
        }
    }
}
