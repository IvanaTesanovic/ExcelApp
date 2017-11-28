using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestApp
{
    public class Summary
    {
        public string Title { get; set; }
        public string Text { get; set; }

        public static List<Summary> GetDummyData()
        {
            List<Summary> dummy = new List<Summary>();
            String lorem = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc congue, est eu molestie pretium, erat augue pellentesque libero, id suscipit ipsum eros quis dui. Nam non orci rutrum eros elementum cursus. Vestibulum nec vehicula tellus. Maecenas rhoncus turpis id mi luctus viverra. Praesent sed nisi eget magna facilisis condimentum a ac ligula. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Pellentesque feugiat, tortor sit amet tristique interdum, erat nibh tempus nisi, sit amet mattis tortor quam ac nibh. Maecenas venenatis erat nec nisl tempor tempus. Pellentesque leo libero, auctor sed tellus ac, gravida aliquam sem. Proin pulvinar erat a lectus ullamcorper, vel luctus enim faucibus. Proin euismod pellentesque elementum. Quisque sit amet finibus quam, eu vehicula justo. Pellentesque bibendum interdum imperdiet.";
            dummy.Add(new Summary() { Title = "Etiology", Text = lorem });
            dummy.Add(new Summary() { Title = "Management and treatment", Text = lorem });
            dummy.Add(new Summary() { Title = "Genetic counseling", Text = lorem });
            dummy.Add(new Summary() { Title = "Prognosis", Text = lorem });

            return dummy;
        }
    }
}
