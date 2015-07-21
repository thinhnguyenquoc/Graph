﻿using AZReport.Model;
using AZReport.Repositories.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Repositories
{
    public class LevelRepository : GenericRepository<Level>, ILevelRepository
    {
        public LevelRepository(AZModelContainer context)
            : base(context)
        {

        }

        public override IEnumerable<Level> GetAll()
        {
            return _entities.Set<Level>().AsEnumerable();
        }
    }
}
