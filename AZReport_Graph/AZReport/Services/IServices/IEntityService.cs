using AZReport.Model;
using AZReport.Repositories.Interface;
using AZReport.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AZReport.Services.IServices
{
    public interface IEntityService<T> where T : AZModel
    {
        void Create(T entity);
        void Delete(T entity);
        IEnumerable<T> GetAll();
        void Update(T entity);
        void Save();
    }
    
}
