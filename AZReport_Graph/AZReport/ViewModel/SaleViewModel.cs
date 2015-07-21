using AZReport.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.ViewModel
{
    public class SaleViewModel
    {
        public string Code { get; set; }
        public List<Sale> Sales { get; set; }
    }
}
