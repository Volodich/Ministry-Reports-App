using System;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using MinistryReports.ViewModels;

namespace MinistryReports.Services
{
    public interface IBakcupService
    {
        void Create(string puth);
        UserSettings GetLoadSettings(string puth);
    }
    public class BackupService : IBakcupService
    {
        private string _puthDefault;
        private string _xmlFormatFile = "*xml";

        private string _rootDir = Path.GetPathRoot(Environment.CurrentDirectory);

        private UserSettings _userSettings;

        public BackupService(UserSettings userSettings)
        {
            _userSettings = userSettings;
            // Укажим шаблонную директорию для поиска файлов настройки
            _puthDefault = Path.Combine(_rootDir, ApplicationConfig.FolderName, ApplicationConfig.SettingsFolder);
        }

        // Возможность использовать путь "по умолчанию"
        public void Create(string puth)
        {
            if (string.IsNullOrEmpty(puth))
            {
                puth = Path.Combine(_puthDefault, "data.xml");
            }

            if (_userSettings == null)
                return;

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(UserSettings));

            using (FileStream fs = new FileStream(puth, FileMode.Create, FileAccess.ReadWrite))
            {
                xmlSerializer.Serialize(fs, _userSettings);
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
