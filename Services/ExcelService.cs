using ClosedXML.Excel;
using ExcelProject.Models;

namespace ExcelProject.Services
{
    public class ExcelService
    {
        private readonly IWebHostEnvironment _env;

        public ExcelService(IWebHostEnvironment env)
        {
            _env = env;
        }

        private string FilePath => Path.Combine(_env.WebRootPath, "uploads", "students.xlsx");

        // ─── FILE CHECK ─────────────────────────────────────────
        private void EnsureFileExists()
        {
            var dir = Path.GetDirectoryName(FilePath)!;
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            if (!File.Exists(FilePath))
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Students");

                ws.Cell(1, 1).Value = "Id";
                ws.Cell(1, 2).Value = "Name";
                ws.Cell(1, 3).Value = "Email";
                ws.Cell(1, 4).Value = "Department";
                ws.Cell(1, 5).Value = "Phone";
                ws.Cell(1, 6).Value = "Status";

                var header = ws.Row(1);
                header.Style.Font.Bold = true;
                header.Style.Fill.BackgroundColor = XLColor.FromHtml("#1e293b");
                header.Style.Font.FontColor = XLColor.White;

                wb.SaveAs(FilePath);
            }
        }

        // ─── GET ALL ───────────────────────────────────────────
        public List<Student> GetStudents()
        {
            EnsureFileExists();

            var list = new List<Student>();

            using var wb = new XLWorkbook(FilePath);
            var ws = wb.Worksheet(1);
            var used = ws.RangeUsed();

            if (used == null) return list;

            foreach (var row in used.RowsUsed().Skip(1))
            {
                list.Add(new Student
                {
                    Id = row.Cell(1).GetValue<int>(),
                    Name = row.Cell(2).GetValue<string>(),
                    Email = row.Cell(3).GetValue<string>(),
                    Department = row.Cell(4).GetValue<string>(),
                    Phone = row.Cell(5).GetValue<string>(),
                    Status = row.Cell(6).GetValue<string>()
                });
            }

            return list;
        }

        // ─── SAVE ALL ──────────────────────────────────────────
        public void SaveStudents(List<Student> students)
        {
            EnsureFileExists();

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Students");

            string[] headers = { "Id", "Name", "Email", "Department", "Phone", "Status" };

            for (int i = 0; i < headers.Length; i++)
                ws.Cell(1, i + 1).Value = headers[i];

            var header = ws.Range(1, 1, 1, headers.Length);
            header.Style.Font.Bold = true;
            header.Style.Fill.BackgroundColor = XLColor.FromHtml("#1e293b");
            header.Style.Font.FontColor = XLColor.White;

            int row = 2;

            foreach (var s in students)
            {
                ws.Cell(row, 1).Value = s.Id;
                ws.Cell(row, 2).Value = s.Name;
                ws.Cell(row, 3).Value = s.Email;
                ws.Cell(row, 4).Value = s.Department;
                ws.Cell(row, 5).Value = s.Phone;
                ws.Cell(row, 6).Value = s.Status;
                row++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(FilePath);
        }

        // ─── DUPLICATE CHECK (🔥 NEW) ──────────────────────────
        private bool IsDuplicate(string email, string phone, int? ignoreId = null)
        {
            var list = GetStudents();

            return list.Any(s =>
                (ignoreId == null || s.Id != ignoreId) &&
                (
                    s.Email.Equals(email, StringComparison.OrdinalIgnoreCase) ||
                    s.Phone == phone
                )
            );
        }

        // ─── GET BY ID ─────────────────────────────────────────
        public Student? GetById(int id)
        {
            return GetStudents().FirstOrDefault(s => s.Id == id);
        }

        // ─── ADD (FIXED) ───────────────────────────────────────
        public bool Add(Student student)
        {
            if (IsDuplicate(student.Email, student.Phone))
                return false;

            var list = GetStudents();

            student.Id = list.Count > 0 ? list.Max(s => s.Id) + 1 : 1;

            list.Add(student);

            SaveStudents(list);

            return true;
        }

        // ─── UPDATE (FIXED) ────────────────────────────────────
        public bool Update(Student updated)
        {
            var list = GetStudents();

            if (IsDuplicate(updated.Email, updated.Phone, updated.Id))
                return false;

            var existing = list.FirstOrDefault(s => s.Id == updated.Id);
            if (existing == null) return false;

            existing.Name = updated.Name;
            existing.Email = updated.Email;
            existing.Department = updated.Department;
            existing.Phone = updated.Phone;
            existing.Status = updated.Status;

            SaveStudents(list);

            return true;
        }

        // ─── DELETE ────────────────────────────────────────────
        public bool Delete(int id)
        {
            var list = GetStudents();

            var item = list.FirstOrDefault(s => s.Id == id);
            if (item == null) return false;

            list.Remove(item);

            SaveStudents(list);

            return true;
        }

        // ─── IMPORT ────────────────────────────────────────────
        public (int added, int skipped) ImportFromFile(IFormFile file)
        {
            var existing = GetStudents();
            var emails = existing.Select(s => s.Email.ToLower()).ToHashSet();
            var phones = existing.Select(s => s.Phone).ToHashSet();

            int added = 0, skipped = 0;

            using var stream = file.OpenReadStream();
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);

            var used = ws.RangeUsed();
            if (used == null) return (0, 0);

            int nextId = existing.Count > 0 ? existing.Max(s => s.Id) + 1 : 1;

            foreach (var row in used.RowsUsed().Skip(1))
            {
                var email = row.Cell(3).GetValue<string>();
                var phone = row.Cell(5).GetValue<string>();

                if (emails.Contains(email.ToLower()) || phones.Contains(phone))
                {
                    skipped++;
                    continue;
                }

                existing.Add(new Student
                {
                    Id = nextId++,
                    Name = row.Cell(2).GetValue<string>(),
                    Email = email,
                    Department = row.Cell(4).GetValue<string>(),
                    Phone = phone,
                    Status = row.Cell(6).GetValue<string>()
                });

                added++;
            }

            SaveStudents(existing);

            return (added, skipped);
        }

        // ─── EXPORT ────────────────────────────────────────────
        public byte[] ExportToBytes()
        {
            EnsureFileExists();
            return File.ReadAllBytes(FilePath);
        }
    }
}