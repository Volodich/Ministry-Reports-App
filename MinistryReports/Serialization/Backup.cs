using System;
using System.IO;
using System.Xml.Serialization;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Windows.Documents;
using System.Windows.Controls;
using MinistryReports;
using MinistryReports.Models;
using MinistryReports.ViewModels;
using MinistryReports.Controllers;
using System.Windows.Forms;

namespace MinistryReports.Serialization
{
    //TODO: async await
    public class Backup
    {
        string puthDefault;
        string XmlFormatFile = "*xml";

        UserSettings userSettings;

        public Backup(UserSettings userSettings)
        {
            this.userSettings = userSettings;
            // Укажим шаблонную директорию для поиска файлов настройки
            var root = Path.GetPathRoot(Environment.CurrentDirectory);
            string programFolder = "Ministry Reports";
            puthDefault = Path.Combine(root, programFolder, "Settings");
        }
        
        // Возможность использовать путь "по умолчанию"
        public bool CreateBackup(string puth = "default")
        {
            if (puth == "default")
            {
                puth = Path.Combine(puthDefault, "data.xml");
            }

            if (userSettings == null)
                return false;
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(UserSettings));

            using (FileStream fs = new FileStream(puth, FileMode.Create, FileAccess.ReadWrite))
            {
                xmlSerializer.Serialize(fs, userSettings);
            }

            return true;
        }

        public bool LoadBackup(out UserSettings userSettings, string puth = "default")
        {

           if (puth == "default")
           {
               puth = Directory.GetFiles(puthDefault, XmlFormatFile)[0];
           }

           XmlSerializer xml = new XmlSerializer(typeof(UserSettings));

           
           using (FileStream fs = new FileStream(puth, FileMode.Open, FileAccess.Read))
           {
               userSettings = (UserSettings) xml.Deserialize(fs);
           }
            
            return true;
        }

        public bool CreateBackupAsync(string puth = "default")
        {
            //TODO: Complete async-await
            return false;
        }
    }
}

