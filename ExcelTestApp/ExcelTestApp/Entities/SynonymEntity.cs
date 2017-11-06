using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp.Entities
{
    public class SynonymEntity
    {
        public Guid Id { get; set; }

        public string Name { get; set; }

        public Guid DiseaseId { get; set; }
    }
}
