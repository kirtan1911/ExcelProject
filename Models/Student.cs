using System.ComponentModel.DataAnnotations;

namespace ExcelProject.Models
{
    public class Student
    {
        public int Id { get; set; }

        // ─── NAME VALIDATION ─────────────────────────────
        [Required(ErrorMessage = "Name is required")]
        [StringLength(50, MinimumLength = 2, ErrorMessage = "Name must be 2–50 characters")]
        [RegularExpression(@"^[a-zA-Z\s]+$", ErrorMessage = "Name can only contain letters and spaces")]
        public string Name { get; set; } = string.Empty;

        // ─── EMAIL VALIDATION ────────────────────────────
        [Required(ErrorMessage = "Email is required")]
        [EmailAddress(ErrorMessage = "Invalid email format")]
        [RegularExpression(
            @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$",
            ErrorMessage = "Invalid email address")]
        public string Email { get; set; } = string.Empty;

        // ─── DEPARTMENT VALIDATION ───────────────────────
        [Required(ErrorMessage = "Department is required")]
        [StringLength(50, MinimumLength = 2)]
        [RegularExpression(@"^[a-zA-Z\s]+$", ErrorMessage = "Department can only contain letters")]
        public string Department { get; set; } = string.Empty;

        // ─── PHONE VALIDATION ────────────────────────────
        [Required(ErrorMessage = "Phone number is required")]
        [RegularExpression(
            @"^[6-9]\d{9}$",
            ErrorMessage = "Phone must be 10 digits and start with 6-9")]
        public string Phone { get; set; } = string.Empty;

        // ─── STATUS VALIDATION ───────────────────────────
        [Required]
        [RegularExpression("Active|Inactive", ErrorMessage = "Status must be Active or Inactive")]
        public string Status { get; set; } = "Active";
    }
}