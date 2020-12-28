using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinistryReports.Models
{
    public class S21Settings
    {
        public string PuthToTamplate { get; set; } // путь к шаблону бланка S-21.

        public string PuthToFolderUnlaoding { get; set; } // путь к папке, в которой создадуться PDF документы всех возвещателей.

        public string PuthToExcelDbFile { get; set; } // путь к "базе данных" - документу Excel, в котором храняться данные возвещателей (дата рождения, крещение, назначение и тп.).
        
        public string NameTable { get; set; } // имя листа, в котором храняться данные возвещателей.
    }
}
