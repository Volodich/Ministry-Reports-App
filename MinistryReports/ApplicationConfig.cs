namespace MinistryReports
{
    public static class ApplicationConfig
    {
        public const string AppName = "MinistryReports"; // название программмы
        public const string FolderName = "MinistryReports"; // название основной папки программы
        public const string SettingsFolder = "Settings"; // название папки с настройками
        public const string DataDir = "ReportsPublishers"; // папка, где лежат все файлы для работы программы
        // s21
        public const string PdfTamplate = "s21blank.pdf"; // названия файла шаблона
               
        public const string FontName = "TimesNewRoman.ttf"; // название шрифта, который используется
        // excel file
        public static class JwExcelBook
        {
            public const string TableName = "reports";
        }

        public static class PublisherInfo
        {
            public const string TableName = "Publishers";
        }
    }
}
