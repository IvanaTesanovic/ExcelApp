using NPoco;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp.Entities
{
    [TableName("Summary")]
    public class SummaryEntity
    {
        public Guid Id { get; set; }

        public string Title { get; set; }

        public string Text { get; set; }

        public Guid DiseaseId { get; set; }
    }
}
