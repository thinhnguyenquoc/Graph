using AZReport.Model;
using AZReport.Repositories.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Services
{
    public class TimeSettingService : EntityService<TimeSetting>, ITimeSettingService
    {
        ITimeSettingRepository _programRepository;
        public TimeSettingService(ITimeSettingRepository countryRepository)
            : base(countryRepository)
        {
            _programRepository = countryRepository;
        }
    }
}
