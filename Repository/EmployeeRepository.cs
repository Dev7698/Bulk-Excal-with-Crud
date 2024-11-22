using CrudUsingAjax.DataBase;
using Microsoft.EntityFrameworkCore;

namespace CrudUsingAjax.Repositories
{
    public class EmployeeRepository<TEntity> : IEmployeeRepository<TEntity> where TEntity : class
    {
        private readonly AppDbContext _context;
        private readonly DbSet<TEntity> _dbSet;

        public EmployeeRepository(AppDbContext context)
        {
            _context = context;
            _dbSet = _context.Set<TEntity>();
        }

        public IQueryable<TEntity> GetAll()
        {
            return _dbSet;
        }

        public TEntity GetById(int id)
        {
            return _dbSet.Find(id);
        }
        public void Add(TEntity entity) 
        {
            _dbSet.Add(entity);
        }

        public void Update(TEntity entity)
        {
            _context.Entry(entity).State = EntityState.Modified;
        }

        public void Delete(int id)
        {
            var entity = _dbSet.Find(id);
            if (entity != null)
            {
                _dbSet.Remove(entity);
            }
        }
        public void AddRange(IEnumerable<TEntity> entities)
        {
            _dbSet.AddRange(entities);
        }
        public void SaveChanges()
        {
            _context.SaveChanges();
        }
    }
}
