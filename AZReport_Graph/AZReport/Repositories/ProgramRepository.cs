using AZReport.Model;
using AZReport.Repositories.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Repositories
{
    public class ProgramRepository : GenericRepository<Program>, IProgramRepository
    {
        public ProgramRepository(AZModelContainer context)
            : base(context)
        {

        }

        public override IEnumerable<Program> GetAll()
        {
            return _entities.Set<Program>().AsEnumerable();
        }
    }
}
