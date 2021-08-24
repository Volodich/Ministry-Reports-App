namespace MinistryReports.Models
{
    class JWMonthReport
    {
        public int Publications { get; set; }
        public int Videos { get; set; }
        public int Hours { get; set; }
        public int ReturnVisits { get; set; }
        public int BibleStudy { get; set; }
        public int CountReports { get; set; }
        public string Notice { get; set; }

        public string Type { get; set; } // тип возвещателя - нужен, если класс используется лоя отчёта всего сорания.
        public string Month {get;set;} // месяц, за который здан отчёт
        public string Year { get; set; } // год за который здан отчёт
        public int CountActivePublishers { get; set; }
    }
}
