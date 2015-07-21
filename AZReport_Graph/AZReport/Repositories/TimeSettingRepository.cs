using AZReport.Model;
using AZReport.Repositories.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Repositories
{
    public class TimeSettingRepository : GenericRepository<TimeSetting>, ITimeSettingRepository
    {
        public TimeSettingRepository(AZModelContainer context)
            : base(context)
        {

        }

        public override IEnumerable<TimeSetting> GetAll()
        {
            return _entities.Set<TimeSetting>().AsEnumerable();
        }
    }
}
