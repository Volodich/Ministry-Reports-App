using System;
using System.Collections;
using MinistryReports.Models;

namespace MinistryReports.ViewModels
{
    class JwMonthReportMeet
    {

        // Сумма отчётов всех возвещателей 
        public JWMonthReport Publishers { get; set; }

        //Сумма отчётов всех подсобных пионеров
        public JWMonthReport AuxiliaryPioneer { get; set; }

        //Сумма отчётов всех пионеров
        public JWMonthReport Pioner { get; set; }

        //Количество присутствующих на собрании
        public int countActivePublisher { get; set; }

        
    }
}
