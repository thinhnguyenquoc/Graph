using AZReport.Model;
using AZReport.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Repositories.Interface
{
    public interface IStoredProcedureRepository : IGenericRepository<AZModel>
    {
        List<ProductivityViewModel> GetProductivity(DateTime start, DateTime end);
        List<ReportViewModel> GetTotalQuantity(DateTime start, DateTime end);
        List<ReportViewModel> GetFreq(DateTime start, DateTime end, DateTime time);
        List<ReportViewModel> GetItemFreq(DateTime start, DateTime end, DateTime time, string Code);
    }
}
