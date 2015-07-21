using AZReport.Model;
using AZReport.Repositories.Interface;
using AZReport.Services.IServices;
using AZReport.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Services
{
    public class SaleService : EntityService<Sale>, ISaleService
    {
        ISaleRepository _programRepository;
        public SaleService(ISaleRepository countryRepository)
            : base(countryRepository)
        {
            _programRepository = countryRepository;
        }

        public Sale CheckAndUpdate(Sale sale)
        {
            var pr = _programRepository.FindBy(x => x.Code == sale.Code && x.Date == sale.Date).FirstOrDefault();
            if (pr == null)
            {
                _programRepository.Add(sale);
            }
            else
            {
                _programRepository.Edit(pr);
            }
            return sale;
        }

        public List<Sale> GetQuantity(DateTime start, DateTime end, String Code)
        {
            return _programRepository.FindBy(x => x.Code == Code && x.Date >= start && x.Date <= end).ToList();
        }
    }
}
