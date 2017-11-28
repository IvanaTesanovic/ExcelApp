using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp
{
    public class TranslationModel
    {
        public string Name { get; private set; }
        private Dictionary<string, string> Words { get; set; }

        public TranslationModel(string name)
        {
            Name = name;
            Words = new Dictionary<string, string>();
        }

        public string[] GetMappedEntities()
        {
            return Words.Keys.ToArray();
        }

        public string[] GetWords()
        {
            return Words.Values.ToArray();
        }

        public void MapTranslation(string original, string translated)
        {
            if (Words.Keys.Contains(original))
            {
                Words[original] = translated;
            }
            else
            {
                Words.Add(original, translated);
            }
        }

        public string GetTranslation(string word)
        {
            if (!Words.Keys.Contains(word))
            {
                return null;
            }
            return Words[word];
        }

    }
}
