using AZReport.Model;
using AZReport.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Services.IServices
{
    public interface ISaleService : IEntityService<Sale>
    {
        Sale CheckAndUpdate(Sale sale);

        List<Sale> GetQuantity(DateTime start, DateTime end, String Code);
    }
}
