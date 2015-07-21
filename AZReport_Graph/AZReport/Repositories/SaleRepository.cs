using AZReport.Model;
using AZReport.Repositories.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Repositories
{
    public class SaleRepository : GenericRepository<Sale>, ISaleRepository
    {
        public SaleRepository(AZModelContainer context)
            : base(context)
        {

        }

        public override IEnumerable<Sale> GetAll()
        {
            return _entities.Set<Sale>().AsEnumerable();
        }
    }
}
