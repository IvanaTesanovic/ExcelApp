using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp
{
    public class Synonym
    {
        public string Name { get; set; }

        public static List<Synonym> getDummyData()
        {
            List<Synonym> dummy = new List<Synonym>();
            dummy.Add(new Synonym() { Name = "Sinonim 1 " });
            dummy.Add(new Synonym() { Name = "Sin 2 " });
            dummy.Add(new Synonym() { Name = "Sino 3 " });
            dummy.Add(new Synonym() { Name = "Sinonim 4 Sinonim " });
            return dummy;
        }
    }
}
