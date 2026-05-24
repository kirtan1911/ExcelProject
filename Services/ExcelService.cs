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

        // ─── Ensure file exists with headers ───────────────────────────────
        private void EnsureFileExists()
        {
            var dir = Path.GetDirectoryName(FilePath)!;
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

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
                // Style header row
                var headerRow = ws.Row(1);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.FromHtml("#1e293b");
                headerRow.Style.Font.FontColor = XLColor.White;
                wb.SaveAs(FilePath);
            }
        }

        // ─── READ ────────────────────────────────────────────────────────────
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
                    Id         = row.Cell(1).GetValue<int>(),
                    Name       = row.Cell(2).GetValue<string>(),
                    Email      = row.Cell(3).GetValue<string>(),
                    Department = row.Cell(4).GetValue<string>(),
                    Phone      = row.Cell(5).GetValue<string>(),
                    Status     = row.Cell(6).GetValue<string>()
                });
            }
            return list;
        }

        // ─── SAVE (overwrite) ────────────────────────────────────────────────
        public void SaveStudents(List<Student> students)
        {
            EnsureFileExists();
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Students");

            // Headers
            string[] headers = { "Id", "Name", "Email", "Department", "Phone", "Status" };
            for (int i = 0; i < headers.Length; i++)
                ws.Cell(1, i + 1).Value = headers[i];

            var headerRange = ws.Range(1, 1, 1, headers.Length);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#1e293b");
            headerRange.Style.Font.FontColor = XLColor.White;

            // Data rows
            int rowIdx = 2;
            foreach (var s in students)
            {
                ws.Cell(rowIdx, 1).Value = s.Id;
                ws.Cell(rowIdx, 2).Value = s.Name;
                ws.Cell(rowIdx, 3).Value = s.Email;
                ws.Cell(rowIdx, 4).Value = s.Department;
                ws.Cell(rowIdx, 5).Value = s.Phone;
                ws.Cell(rowIdx, 6).Value = s.Status;
                rowIdx++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(FilePath);
        }

        // ─── GET BY ID ───────────────────────────────────────────────────────
        public Student? GetById(int id) => GetStudents().FirstOrDefault(s => s.Id == id);

        // ─── ADD ─────────────────────────────────────────────────────────────
        public void Add(Student student)
        {
            var list = GetStudents();
            student.Id = list.Count > 0 ? list.Max(s => s.Id) + 1 : 1;
            list.Add(student);
            SaveStudents(list);
        }

        // ─── UPDATE ──────────────────────────────────────────────────────────
        public bool Update(Student updated)
        {
            var list = GetStudents();
            var existing = list.FirstOrDefault(s => s.Id == updated.Id);
            if (existing == null) return false;

            existing.Name       = updated.Name;
            existing.Email      = updated.Email;
            existing.Department = updated.Department;
            existing.Phone      = updated.Phone;
            existing.Status     = updated.Status;

            SaveStudents(list);
            return true;
        }

        // ─── DELETE ──────────────────────────────────────────────────────────
        public bool Delete(int id)
        {
            var list = GetStudents();
            var item = list.FirstOrDefault(s => s.Id == id);
            if (item == null) return false;
            list.Remove(item);
            SaveStudents(list);
            return true;
        }

        // ─── IMPORT FROM UPLOADED FILE ───────────────────────────────────────
        public (int added, int skipped) ImportFromFile(IFormFile file)
        {
            var existing = GetStudents();
            var existingEmails = existing.Select(s => s.Email.ToLower()).ToHashSet();
            int added = 0, skipped = 0;

            using var stream = file.OpenReadStream();
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet(1);
            var used = ws.RangeUsed();
            if (used == null) return (0, 0);

            int nextId = existing.Count > 0 ? existing.Max(s => s.Id) + 1 : 1;

            foreach (var row in used.RowsUsed().Skip(1))
            {
                var email = row.Cell(3).GetValue<string>().ToLower();
                if (existingEmails.Contains(email)) { skipped++; continue; }

                existing.Add(new Student
                {
                    Id         = nextId++,
                    Name       = row.Cell(2).GetValue<string>(),
                    Email      = row.Cell(3).GetValue<string>(),
                    Department = row.Cell(4).GetValue<string>(),
                    Phone      = row.Cell(5).GetValue<string>(),
                    Status     = row.Cell(6).GetValue<string>()
                });
                added++;
            }

            SaveStudents(existing);
            return (added, skipped);
        }

        // ─── EXPORT (returns file bytes) ─────────────────────────────────────
        public byte[] ExportToBytes()
        {
            EnsureFileExists();
            return File.ReadAllBytes(FilePath);
        }
    }
}
