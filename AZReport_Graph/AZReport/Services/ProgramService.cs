using AZReport.Model;
using AZReport.Repositories.Interface;
using AZReport.Services.IServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Services
{
    public class ProgramService : EntityService<Program>, IProgramService
    {
        IProgramRepository _programRepository;

        public ProgramService(IProgramRepository countryRepository)
            : base(countryRepository)
        {
            _programRepository = countryRepository;
        }

        public Program InsertOrUpdate(Program program)
        {
            if (program.Id != 0 && program.Id != null)
            {
                _programRepository.Add(program);
            }
            else
            {
                _programRepository.Edit(program);
            }
            return program;
        }

        public Program CheckAndUpdate(Program program)
        {
            var pr = _programRepository.FindBy(x => x.Code == program.Code).FirstOrDefault();
            if (pr == null)
            {
                _programRepository.Add(program);
            }
            else
            {
                _programRepository.Edit(pr);
            }
            return program;
        }

        public Program CheckAndCreate(Program program)
        {
            var pr = _programRepository.FindBy(x => x.Code == program.Code).FirstOrDefault();
            if (pr == null)
            {
                _programRepository.Add(program);
            }
            else
            {
                //_programRepository.Edit(pr);
            }
            return program;
        }

        public Program FindItem(string code)
        {
            return _programRepository.FindBy(x => x.Code == code).FirstOrDefault();
        }
    }
}
