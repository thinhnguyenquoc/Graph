using AZReport.Model;
using AZReport.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Services.IServices
{
    public interface IScheduleService : IEntityService<Schedule>
    {
        Schedule CheckAndCreate(Schedule schedule);
        List<Schedule> GetByDate(DateTime start, DateTime end);
        bool CheckExistDate(DateTime date);
        void DeleteOldDate(DateTime date);
    }
}
