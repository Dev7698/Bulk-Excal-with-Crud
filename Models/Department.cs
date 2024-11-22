using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace CrudUsingAjax.Models
{
    public class Department
    {
        [Key]
        public int Department_Id { get; set; }
        [EnumDataType(typeof(DepartmentName))]  
        public DepartmentName? DepartmentName { get; set; }
    }
    public enum DepartmentName
    {
        [Description("HR")]
        HR = 1,
        [Description("Developer")]
        Developer = 2,
        [Description("Finance")]
        Finance = 3,
        [Description("Marketing")]
        Marketing = 4
    }

}
