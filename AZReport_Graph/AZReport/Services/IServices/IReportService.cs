using AZReport.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Services.IServices
{
    public interface IReportService
    {
        List<ProductivityViewModel> GetProductivity(DateTime start, DateTime end);
        List<ReportViewModel> GetQuantity(DateTime start, DateTime end);
        List<ReportViewModel> GetFreq(DateTime start, DateTime end, DateTime time);
        List<ReportViewModel> GetItemFreq(DateTime start, DateTime end, DateTime time, string code);
    }
}
