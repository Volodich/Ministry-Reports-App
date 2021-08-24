namespace MinistryReports
{
    public static class ApplicationConfig
    {
        public static string AppName = "MinistryReports"; // название программмы
        public static string FolderName = "MinistryReports"; // название основной папки программы
        public static string SettingsFolder = "Settings"; // название папки с настройками
        public static string DataDir = "ReportsPublishers"; // папка, где лежат все файлы для работы программы
        // s21
        public static string PdfTamplate = "s21blank.pdf"; // названия файла шаблона

        public static string FontName = "TimesNewRoman.ttf"; // название шрифта, который используется
        // excel file
        public static class JwExcelBook
        {
            public static string TableName = "reports";
        } 
    }
}
