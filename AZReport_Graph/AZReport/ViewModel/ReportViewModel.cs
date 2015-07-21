using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.ViewModel
{
    public class ReportViewModel
    {
        public string Code { get; set; }
        public int? Quantity { get; set; }
        public DateTime? Date { get; set; }
        public int Freq { get; set; }
    }
}
