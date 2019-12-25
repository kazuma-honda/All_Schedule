using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Schedule
{
    public class ScheduleItem
    {
        public string Title { get; set; }

        public DateTime Start { get; set; }

        public DateTime End { get; set; }

        public ScheduleItem(string title, DateTime start, DateTime end)
        {
            // this.Title = title;
            // this.Start = start;
            // this.End = end;         
            
            this.Title = title;
            this.Start = start;
            this.End = end;
        }

        public override string ToString()
        {
            //return $"({Start}～{End}) {Title}";
            return $"({Start.ToString("MM/dd HH:mm")}～{End.ToString("HH:mm")}) {Title}";
        }
    }
}
