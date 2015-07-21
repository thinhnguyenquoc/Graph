using AZReport.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.ViewModel
{
    public class ScheduleViewModel
    {
        public int Hour { get; set; }
        public List<Schedule> programs { get; set; }
    }
}
