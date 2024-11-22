namespace CrudUsingAjax.Repositories
{
    public interface IEmployeeRepository<TEntity> where TEntity : class
    {
        IQueryable<TEntity> GetAll();
        TEntity GetById(int id);
        void Add(TEntity entity);
        void Update(TEntity entity);
        void Delete(int id);
        void AddRange(IEnumerable<TEntity> entities);
        void SaveChanges();
    }
}
