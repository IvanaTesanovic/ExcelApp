using ExcelTestApp.Entities;
using System.Collections.Generic;

namespace ExcelTestApp
{
    public class Summary
    {
        public string Title { get; set; }
        public string Text { get; set; }

        public Summary(SummaryEntity summary)
        {
            Title = summary.Title;
            Text = summary.Text;
        }
    }
}
