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

        // ─── INDEX ────────────────────────────────────────────────────────────
        public IActionResult Index(string? search, string? status)
        {
            var list = _excel.GetStudents();

            if (!string.IsNullOrWhiteSpace(search))
                list = list.Where(s =>
                    s.Name.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                    s.Email.Contains(search, StringComparison.OrdinalIgnoreCase) ||
                    s.Department.Contains(search, StringComparison.OrdinalIgnoreCase)).ToList();

            if (!string.IsNullOrWhiteSpace(status))
                list = list.Where(s => s.Status == status).ToList();

            ViewBag.Search = search;
            ViewBag.Status = status;
            ViewBag.Total  = _excel.GetStudents().Count;
            return View(list);
        }

        // ─── ADD ──────────────────────────────────────────────────────────────
        [HttpPost]
        public IActionResult Add(Student student)
        {
            if (!ModelState.IsValid)
                return RedirectToAction(nameof(Index));

            _excel.Add(student);
            TempData["Success"] = $"Student \"{student.Name}\" added successfully!";
            return RedirectToAction(nameof(Index));
        }

        // ─── EDIT GET ─────────────────────────────────────────────────────────
        public IActionResult Edit(int id)
        {
            var student = _excel.GetById(id);
            if (student == null) return NotFound();
            return Json(student);
        }

        // ─── EDIT POST ────────────────────────────────────────────────────────
        [HttpPost]
        public IActionResult EditPost(Student student)
        {
            var ok = _excel.Update(student);
            TempData[ok ? "Success" : "Error"] = ok
                ? $"Student \"{student.Name}\" updated successfully!"
                : "Student not found.";
            return RedirectToAction(nameof(Index));
        }

        // ─── DELETE ───────────────────────────────────────────────────────────
        [HttpPost]
        public IActionResult Delete(int id)
        {
            var student = _excel.GetById(id);
            var ok = _excel.Delete(id);
            TempData[ok ? "Success" : "Error"] = ok
                ? $"Student deleted successfully."
                : "Student not found.";
            return RedirectToAction(nameof(Index));
        }

        // ─── UPLOAD VIEW ──────────────────────────────────────────────────────
        public IActionResult Upload() => View();

        // ─── UPLOAD POST ──────────────────────────────────────────────────────
        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Error"] = "Please select a valid Excel file.";
                return View();
            }

            if (!Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                TempData["Error"] = "Only .xlsx files are supported.";
                return View();
            }

            var (added, skipped) = _excel.ImportFromFile(file);
            TempData["Success"] = $"Import complete! {added} added, {skipped} skipped (duplicates).";
            return RedirectToAction(nameof(Index));
        }

        // ─── EXPORT ───────────────────────────────────────────────────────────
        public IActionResult Export()
        {
            var bytes = _excel.ExportToBytes();
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"students_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
        }
    }
}
