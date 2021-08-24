using System;
using System.IO;
using MinistryReports.Models;
using MinistryReports.ViewModels;

namespace MinistryReports.Services
{
    public interface IUserService
    {
        UserSettings GetUserSettings();
    }
    public class UserService : IUserService
    {
        /// <summary>
        /// Generated base settings for app
        /// </summary>
        /// <returns>base settings (get from ApplicationConfig</returns>
        public UserSettings GetUserSettings()
        {
            return new UserSettings()
            {
                UserName = "admin",
                JWBookSettings = new JWBookSettings()
                {
                    JWBookPuth = Path.Combine(Path.GetPathRoot(Environment.CurrentDirectory), ApplicationConfig.FolderName, ApplicationConfig.SettingsFolder),
                    // TODO: initialize
                },
                S21Settings = new S21Settings()
                {
                    NameTable = ApplicationConfig.JwExcelBook.TableName,
                    PuthToExcelDbFile = Path.Combine(Path.GetPathRoot(Environment.CurrentDirectory), ApplicationConfig.FolderName, ApplicationConfig.SettingsFolder),
                    PuthToFolderUnlaoding = Path.Combine(Path.GetPathRoot(Environment.CurrentDirectory), ApplicationConfig.FolderName, ApplicationConfig.DataDir),
                    PuthToTamplate = Path.Combine(Path.GetPathRoot(Environment.CurrentDirectory), ApplicationConfig.FolderName, ApplicationConfig.SettingsFolder, ApplicationConfig.PdfTamplate)
                }
            };
        }

    }
}
