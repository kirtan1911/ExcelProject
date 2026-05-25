using ExcelProject.Models;
using ExcelProject.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelProject.Controllers
{
    public class StudentController : Controller
    {
        private readonly ExcelService _excel;

        public StudentController(ExcelService excel)
        {
            _excel = excel;
        }

        // ─── INDEX ─────────────────────────────────────────────
        public IActionResult Index(string? search, string? status)
        {
            var students = _excel.GetStudents();

            if (!string.IsNullOrWhiteSpace(search))
            {
                search = search.Trim();

                students = students.Where(s =>
                    (!string.IsNullOrEmpty(s.Name) && s.Name.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(s.Email) && s.Email.Contains(search, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(s.Department) && s.Department.Contains(search, StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }

            if (!string.IsNullOrWhiteSpace(status))
            {
                students = students
                    .Where(s => s.Status != null && s.Status.Equals(status, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            ViewBag.Search = search;
            ViewBag.Status = status;
            ViewBag.Total = _excel.GetStudents().Count;

            return View(students);
        }

        // ─── ADD ───────────────────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Add(Student student)
        {
            if (!ModelState.IsValid)
                return RedirectToAction(nameof(Index));

            _excel.Add(student);

            TempData["Success"] = $"Student \"{student.Name}\" added successfully!";
            return RedirectToAction(nameof(Index));
        }

        // ─── GET SINGLE STUDENT (FOR EDIT MODAL) ───────────────
        [HttpGet]
        public IActionResult GetStudent(int id)
        {
            var student = _excel.GetById(id);
            if (student == null) return NotFound();

            return Json(student);
        }

        // ─── EDIT POST (FIXED) ────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(Student student)
        {
            if (student == null || student.Id <= 0)
            {
                TempData["Error"] = "Invalid student data.";
                return RedirectToAction(nameof(Index));
            }

            var existing = _excel.GetById(student.Id);

            if (existing == null)
            {
                TempData["Error"] = "Student not found.";
                return RedirectToAction(nameof(Index));
            }

            existing.Name = student.Name;
            existing.Email = student.Email;
            existing.Department = student.Department;
            existing.Phone = student.Phone;
            existing.Status = student.Status;

            var ok = _excel.Update(existing);

            TempData[ok ? "Success" : "Error"] = ok
                ? $"Student \"{student.Name}\" updated successfully!"
                : "Update failed.";

            return RedirectToAction(nameof(Index));
        }

        // ─── DELETE ────────────────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Delete(int id)
        {
            var ok = _excel.Delete(id);

            TempData[ok ? "Success" : "Error"] = ok
                ? "Student deleted successfully."
                : "Student not found.";

            return RedirectToAction(nameof(Index));
        }

        // ─── UPLOAD ────────────────────────────────────────────
        public IActionResult Upload() => View();

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Error"] = "Please select a valid Excel file.";
                return View();
            }

            if (!Path.GetExtension(file.FileName)
                .Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                TempData["Error"] = "Only .xlsx files are supported.";
                return View();
            }

            var (added, skipped) = _excel.ImportFromFile(file);

            TempData["Success"] =
                $"Import complete! {added} added, {skipped} skipped.";

            return RedirectToAction(nameof(Index));
        }

        // ─── EXPORT ────────────────────────────────────────────
        public IActionResult Export()
        {
            var bytes = _excel.ExportToBytes();

            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"students_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
        }
    }
}