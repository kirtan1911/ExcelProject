namespace ExcelProject.Helpers
{
    public static class FileHelper
    {
        public static string GetExcelPath(IWebHostEnvironment env)
        {
            var uploadsPath = Path.Combine(env.WebRootPath, "uploads");
            if (!Directory.Exists(uploadsPath))
                Directory.CreateDirectory(uploadsPath);
            return Path.Combine(uploadsPath, "students.xlsx");
        }

        public static bool ExcelExists(IWebHostEnvironment env)
        {
            return File.Exists(GetExcelPath(env));
        }
    }
}
