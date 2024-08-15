
namespace StaffManagement.Models
{
    public class Staff
    {
        public string StaffId { get; set; } // 8 characters
        public string FullName { get; set; } // 100 characters
        public DateTime Birthday { get; set; }
        public int Gender { get; set; } // 1: Male, 2: Female
    }
}
