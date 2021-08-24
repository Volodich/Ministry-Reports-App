using System;
using System.IO;
using System.Reflection;

namespace MinistryReports.Services
{
    public interface IDirectoryService
    {
        void CreateProgramDirectory();
        void CreateSystemFile();
    }
    public class DirectoryService : IDirectoryService
    {
        private readonly string _rootDir;
        public DirectoryService()
        {
            _rootDir = Path.GetPathRoot(Environment.CurrentDirectory);
        }
        public void CreateProgramDirectory()
        {
            Directory.CreateDirectory(Path.Combine(_rootDir, ApplicationConfig.FolderName)); // /BaseDir
            Directory.CreateDirectory(Path.Combine(_rootDir, ApplicationConfig.FolderName, ApplicationConfig.SettingsFolder)); // /BaseDir/Settings
            Directory.CreateDirectory(Path.Combine(_rootDir, ApplicationConfig.FolderName, ApplicationConfig.DataDir)); // BaseDir/Data
        }

        public void CreateSystemFile()
        {
            string appFolderPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            string resourcesFolderPath = Path.Combine(Directory.GetParent(appFolderPath).Parent.FullName, "Resources");
            
            // s21 blank
            File.Copy(
                Path.Combine(resourcesFolderPath, "Pdf", "s21blank.pdf"),
                Path.Combine(_rootDir, ApplicationConfig.FolderName, ApplicationConfig.SettingsFolder, ApplicationConfig.PdfTamplate),
                false);
            // fonts
            File.Copy(
                Path.Combine(resourcesFolderPath, "Fonts", "Times New Roman.ttf"), 
                Path.Combine(_rootDir, ApplicationConfig.FolderName, ApplicationConfig.SettingsFolder, ApplicationConfig.FontName), 
                false);
        }
    }
}
