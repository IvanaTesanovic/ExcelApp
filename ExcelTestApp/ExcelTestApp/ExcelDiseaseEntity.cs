using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp
{
    public class ExcelDiseaseEntity
    {
        public int RowIndex { get; set; }
        public string PropertyName { get; set; }
        public string PropertyValue { get; set; }
        public int NumberOfRows { get; set; }
    }
}
