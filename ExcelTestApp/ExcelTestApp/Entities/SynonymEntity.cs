using NPoco;
using System;

namespace ExcelTestApp.Entities
{
    [TableName("Synonym")]
    public class SynonymEntity
    {
        public Guid Id { get; set; }

        public string Name { get; set; }

        public Guid DiseaseId { get; set; }
    }
}
