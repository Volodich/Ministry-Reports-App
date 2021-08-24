using System;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using MinistryReports.ViewModels;

namespace MinistryReports.Services
{
    public interface IBackupService
    {
        void Create(UserSettings settings, string puth = null);
        UserSettings GetLoadSettings(string puth);
    }
    public class BackupService : IBackupService
    {
        private string _puthDefault;
        private string _xmlFormatFile = "*xml";

        private string _rootDir = Path.GetPathRoot(Environment.CurrentDirectory);

        public BackupService()
        {
            // Укажим шаблонную директорию для поиска файлов настройки
            _puthDefault = Path.Combine(_rootDir, ApplicationConfig.FolderName, ApplicationConfig.SettingsFolder);
        }

        // Возможность использовать путь "по умолчанию"
        public void Create(UserSettings settings, string puth = null)
        {
            if (string.IsNullOrEmpty(puth))
            {
                puth = Path.Combine(_puthDefault, "data.xml");
            }

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(UserSettings));

            using (FileStream fs = new FileStream(puth, FileMode.Create, FileAccess.ReadWrite))
            {
                xmlSerializer.Serialize(fs, settings);
            }
        }

        public UserSettings GetLoadSettings(string puth)
        {
            UserSettings userSettings;

            if (string.IsNullOrEmpty(puth))
            {
                puth = Directory.GetFiles(_puthDefault, _xmlFormatFile).First();
            }

            XmlSerializer xml = new XmlSerializer(typeof(UserSettings));


            using (FileStream fs = new FileStream(puth, FileMode.Open, FileAccess.Read))
            {
                userSettings = (UserSettings)xml.Deserialize(fs);
            }

            return userSettings;
        }
    }
}
