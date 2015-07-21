using AZReport.Model;
using AZReport.Repositories.Interface;
using AZReport.ViewModel;
using System;
using System.Collections.Generic;
using System.Data.Objects;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace AZReport.Repositories
{
    public class StoredProcedureRepository : GenericRepository<AZModel>, IStoredProcedureRepository
    {
        private AZModelContainer _context;
        public StoredProcedureRepository(AZModelContainer context)
            : base(context)
        {
            _context = context;
        }

        public List<ProductivityViewModel> GetProductivity(DateTime start, DateTime end)
        {
            //var startTime = new SqlParameter
            //{
            //    ParameterName = "param1",
            //    Value = start.ToString()
            //};
            //var endTime = new SqlParameter
            //{
            //    ParameterName = "param2",
            //    Value = end.ToString()
            //};
            //var courseList = _context.Database.SqlQuery<ProductivityViewModel>("GetProductivity @param1, @param2 ", startTime, endTime).ToList<ProductivityViewModel>();

             //select distinct p.Code, p.Name, p.Note, p.Category, p.Duration, p.Price 
             //                                   from Schedules s join Programs p on s.Code = p.Code
             //                                   where s.Date between @param1 and @param2

            var query = from sche in _context.Schedules join pro in _context.Programs on sche.Code equals pro.Code
                        where sche.Date >= start && sche.Date <= end
                        select new {
                            Code = sche.Code,
                            Name = pro.Name,
                            Category = pro.Category,
                            Price = pro.Price,
                            Note = pro.Note,
                            Duration = pro.Duration
                        };
            List<ProductivityViewModel> result = new List<ProductivityViewModel>();
            foreach (var i in query.Distinct().ToList())
            {
                var k = new ProductivityViewModel();
                k.Code = i.Code;
                k.Name = i.Name;
                k.Category = i.Category;
                k.Price = i.Price;
                k.Note = i.Note;
                k.Duration = i.Duration;
                result.Add(k);
            }
            return result;
        }

        public List<ReportViewModel> GetTotalQuantity(DateTime start, DateTime end)
        {
            //var startTime = new SqlParameter
            //{
            //    ParameterName = "param1",
            //    Value = start.ToString()
            //};
            //var endTime = new SqlParameter
            //{
            //    ParameterName = "param2",
            //    Value = end.ToString()
            //};
            //var courseList = _context.Database.SqlQuery<ReportViewModel>("GetTotalQuantity @param1, @param2 ", startTime, endTime).ToList<ReportViewModel>();

            //CREATE PROCEDURE [dbo].[GetTotalQuantity]
            //    @param1 datetime ,
            //    @param2 datetime
            //AS
            //    Declare @program Table(
            //        code varchar(max) not null
            //    ) 
            //    insert into @program select distinct p.Code 
            //    from Schedules s join Programs p on s.Code = p.Code
            //    where s.Date between @param1 and @param2
            //    --select * from @program
            //    Declare @quantity Table(
            //        code varchar(max) not null,
            //        quantity int
            //    )
            //    insert into @quantity select s.Code, s.Quantity
            //    --from Sales s join @program p on s.Code = p.code 
            //    from Sales s
            //    where s.Date between @param1 and @param2 
            //    select q.code as Code, Sum(q.quantity) as Quantity from @quantity q
            //    group by q.code
            //RETURN 0 ";

            var codeList = from s in _context.Schedules  join p in _context.Programs  on s.Code equals p.Code
                           where s.Date >= start && s.Date <= end
                           select new {
                               s.Code
                           };
            var quantityList = from s in _context.Sales 
                               where s.Date >= start && s.Date <= end
                               select new {
                                   s.Code,
                                   Quantity = s.Quantity
                               };
            var query = from q in quantityList
                        group q by new {q.Code}
                        into grp
                        select new {
                            grp.Key.Code,
                            Quantity = grp.Sum(t =>t.Quantity)

                        };

            List<ReportViewModel> result = new List<ReportViewModel>();
            foreach (var i in query.Distinct().ToList())
            {
                var k = new ReportViewModel();
                k.Code = i.Code;
                k.Quantity = i.Quantity;
                result.Add(k);
            }
            return result;
        }

        public List<ReportViewModel> GetFreq(DateTime start, DateTime end, DateTime time)
        {
            //var startTime = new SqlParameter
            //{
            //    ParameterName = "param1",
            //    Value = start.ToString()
            //};
            //var endTime = new SqlParameter
            //{
            //    ParameterName = "param2",
            //    Value = end.ToString()
            //};
            //var courseList = _context.Database.SqlQuery<ReportViewModel>("GetTotalFrequency @param1, @param2 ", startTime, endTime).ToList<ReportViewModel>();

            //CREATE PROCEDURE [dbo].[GetTotalFrequency]
            //    @param1 datetime ,
            //    @param2 datetime
            //AS
            //    Declare @temp Table(
            //        Code varchar(max) not null,
            //        Date date
            //    ) 
            //    insert into @temp select s.Code, CONVERT(date, s.Date) 
            //    from Schedules s join Programs p on s.Code = p.Code
            //    where s.Date between @param1 and @param2

            //    select t.Code, t.Date, COUNT(t.Code) Freq
            //    from @temp t
            //    group by t.Code, t.Date
            //RETURN ";

            var codeList = from s in _context.Schedules
                           join p in _context.Programs on s.Code equals p.Code
                           where s.Date >= start && s.Date <= end && s.Date.Hour >= time.Hour && s.Date.Minute >= time.Minute && s.Date.Second >= time.Second
                           select new
                           {
                               Code = s.Code,
                               Date = s.Date
                           };
            var query = from t in codeList
                        group t by new { t.Code, t.Date.Year, t.Date.Month, t.Date.Day }
                            into grp
                            select new
                            {
                                Code = grp.Key.Code,
                                Year = grp.Key.Year,
                                Month = grp.Key.Month, 
                                Day = grp.Key.Day,
                                Freq = grp.Count()
                            };
            List<ReportViewModel> result = new List<ReportViewModel>();
            foreach (var i in query.Distinct().ToList())
            {
                var k = new ReportViewModel();
                k.Code = i.Code;
                k.Date = new DateTime(i.Year, i.Month, i.Day, 0, 0, 0);
                k.Freq = i.Freq;
                result.Add(k);
            }
            return result;
        }

        public List<ReportViewModel> GetItemFreq(DateTime start, DateTime end, DateTime time, string Code)
        {
           
            var codeList = from s in _context.Schedules
                           join p in _context.Programs on s.Code equals p.Code
                           where s.Date >= start && s.Date <= end && s.Date.Hour >= time.Hour && s.Date.Minute >= time.Minute && s.Date.Second >= time.Second && s.Code == Code
                           select new
                           {
                               Code = s.Code,
                               Date = s.Date
                           };
            var query = from t in codeList
                        group t by new { t.Code, t.Date.Year, t.Date.Month, t.Date.Day }
                            into grp
                            select new
                            {
                                Code = grp.Key.Code,
                                Year = grp.Key.Year,
                                Month = grp.Key.Month,
                                Day = grp.Key.Day,
                                Freq = grp.Count()
                            };
            List<ReportViewModel> result = new List<ReportViewModel>();
            foreach (var i in query.Distinct().ToList())
            {
                var k = new ReportViewModel();
                k.Code = i.Code;
                k.Date = new DateTime(i.Year, i.Month, i.Day, 0, 0, 0);
                k.Freq = i.Freq;
                result.Add(k);
            }
            return result;
        }
    }
}
