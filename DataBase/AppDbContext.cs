using CrudUsingAjax.Models;
using Microsoft.EntityFrameworkCore;
namespace CrudUsingAjax.DataBase
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
        {

        }
        public DbSet<Employee> EmployeesDB { get; set; }
        public DbSet<Department> DepartmentDB { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder) 
        {
            modelBuilder.Entity<Department>().HasData(
                new Department { Department_Id = 1, DepartmentName = DepartmentName.HR },  
                new Department { Department_Id = 2, DepartmentName = DepartmentName.Developer }, 
                new Department { Department_Id = 3, DepartmentName = DepartmentName.Finance },
                new Department { Department_Id = 4, DepartmentName = DepartmentName.Marketing }
            );

            modelBuilder.Entity<Employee>()
                .HasOne(e => e.Department)
                .WithMany()
                .HasForeignKey(e => e.Department_Id)
                .OnDelete(DeleteBehavior.Restrict);

            base.OnModelCreating(modelBuilder);
        }
    }
    
}
