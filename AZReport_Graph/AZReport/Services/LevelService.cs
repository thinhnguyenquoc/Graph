using AZReport.Model;
using AZReport.Repositories.Interface;
using AZReport.Services.IServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Services
{
    public class LevelService : EntityService<Level>, ILevelService
    {
        ILevelRepository _levelRepository;

        public LevelService(ILevelRepository countryRepository)
            : base(countryRepository)
        {
            _levelRepository = countryRepository;
        }

    }
}
