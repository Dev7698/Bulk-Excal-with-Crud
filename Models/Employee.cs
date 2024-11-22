using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace CrudUsingAjax.Models
{
    public class Employee
    {
        [Key]
        public int Employee_Id { get; set; }
        [DisplayName("First Name")]
        public string? First_Name { get; set; } 
        [DisplayName("Last Name")]
        public string? Last_Name { get; set; } 
        [DisplayName("Email Id")]
        public string? Email { get; set; }
        [DisplayName("Phone Number")]
        public string? Phone_Number { get; set; }
        [DisplayName("Gender")]
        public string? Gender { get; set; }
        [DisplayName("Department")]
        public int Department_Id { get; set; }
        public Department? Department { get; set; }
        [DisplayName("Joining Date")]
        public DateOnly? Joining_Date { get; set; }
        [DisplayName("Address")]
        public string? Address { get; set; }
        [DisplayName("Profile Image")]
        public string? Profile_Image { get; set; }
    }
}
