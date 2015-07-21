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
   public abstract class EntityService<T> : IEntityService<T> where T : AZModel
    {
        IGenericRepository<T> _repository;

        public EntityService(IGenericRepository<T> repository)
        {
            _repository = repository;
        }


        public virtual void Create(T entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }
            _repository.Add(entity);
        }


        public virtual void Update(T entity)
        {
            if (entity == null) throw new ArgumentNullException("entity");
            _repository.Edit(entity);
        }

        public virtual void Delete(T entity)
        {
            if (entity == null) throw new ArgumentNullException("entity");
            _repository.Delete(entity);
        }

        public virtual void Save()
        {
            _repository.Save();
        }

        public virtual IEnumerable<T> GetAll()
        {
            return _repository.GetAll();
        }
    }
}
