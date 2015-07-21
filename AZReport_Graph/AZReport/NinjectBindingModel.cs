using AZReport.Model;
using AZReport.Repositories;
using AZReport.Repositories.Interface;
using AZReport.Services;
using AZReport.Services.IServices;
using Ninject.Modules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport
{
    public class NinjectBindingModel : NinjectModule
    {
        public override void Load()
        {
            Bind<AZModelContainer>().ToSelf().InSingletonScope();

            Bind<IProgramRepository>().To<ProgramRepository>();
            Bind<IScheduleRepository>().To<ScheduleRepository>();
            Bind<ISaleRepository>().To<SaleRepository>();
            Bind<IStoredProcedureRepository>().To<StoredProcedureRepository>();
            Bind<ILevelRepository>().To<LevelRepository>();
            Bind<ITimeSettingRepository>().To<TimeSettingRepository>();

            Bind<IProgramService>().To<ProgramService>();
            Bind<IScheduleService>().To<ScheduleService>();
            Bind<ISaleService>().To<SaleService>();
            Bind<IReportService>().To<ReportService>();
            Bind<ILevelService>().To<LevelService>();
            Bind<ITimeSettingService>().To<TimeSettingService>();
        }
    }
}
