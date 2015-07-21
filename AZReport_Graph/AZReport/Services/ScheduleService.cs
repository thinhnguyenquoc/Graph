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
    public class ScheduleService : EntityService<Schedule>, IScheduleService
    {
        IScheduleRepository _scheduleRepository;

        public ScheduleService(IScheduleRepository countryRepository)
            : base(countryRepository)
        {
            _scheduleRepository = countryRepository;
        }

        public Schedule CheckAndCreate(Schedule schedule)
        {
            var sche = _scheduleRepository.FindBy(x => x.Code == schedule.Code && x.Date == schedule.Date).FirstOrDefault();
            if (sche == null)
            {
                _scheduleRepository.Add(schedule);
            }
            return schedule;
        }

        public List<Schedule> GetByDate(DateTime start, DateTime end)
        {
            DateTime startTime = new DateTime(start.Year, start.Month, start.Day, 0,0,0);
            DateTime endTime = new DateTime(end.Year, end.Month, end.Day, 23, 59, 59);
            return _scheduleRepository.FindBy(x => x.Date <= endTime && x.Date >= startTime).ToList();
        }

        public bool CheckExistDate(DateTime date)
        {
            var result = _scheduleRepository.FindBy(x => x.Date.Year == date.Year && x.Date.Month == date.Month && x.Date.Day == date.Day).FirstOrDefault();
            return result != null ? true : false;
        }

        public void DeleteOldDate(DateTime date)
        {
            var programs = _scheduleRepository.FindBy(x => x.Date.Year == date.Year && x.Date.Month == date.Month && x.Date.Day == date.Day).ToList();
            foreach(var i in programs){
                _scheduleRepository.Delete(i);
            }
            _scheduleRepository.Save();
        }
    }
}
