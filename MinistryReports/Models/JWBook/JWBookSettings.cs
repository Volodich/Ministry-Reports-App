namespace MinistryReports.Models
{
    public class JWBookSettings
    {
        public string ApplicationName => "Ministry Reports";

        public string SpreadsheetId { get; set; }

        public string SheetNameDataMinistry { get; set; } // Имя листа - данные возвещателей о проповедническом служении.

        public string SheetNameArchiveReports { get; set; } // Имя листа - данные, отправляемые в филиал (месяцный отчёт).

        public string SheetNameNoActivityPublishers { get; set; } // Имя листа - данные о неактивных возвещателях.

        public string PuthJsonFile { get; set; } // путь к файлу JSON (для работы с google sheets).

        public int SheetIdArchiveReport { get; set; } // Именно этот тип!

        public int SheetIdNoActivityPublisher { get; set; } // Именно этот тип!

        public int SheetIdDataMinistry { get; set; } // Именно этот тип!

        public string JWBookPath { get; set; }

    }
}
