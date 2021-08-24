using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using ConsoleApp4.JWBook;
using MinistryReports.Extensions;
using MinistryReports.Services;
using MinistryReports.ViewModels;

namespace MinistryReports
{
    public class Initialize
    {
        private readonly IBackupService _backupService;
        private readonly IDirectoryService _directoryService;

        private UserSettings _userSettings;

        private MainWindow _mainWindow;
        private Window _waitWindow;

        public Initialize(MainWindow window, Window waitWindow)
        {
            _mainWindow = window;
            _waitWindow = waitWindow;

            _directoryService = new DirectoryService();
            _backupService = new BackupService();
        }

        public async Task<UserSettings> LoadAppSettingsAsync() => await Task.Run(LoadAppSettings);

        public UserSettings LoadAppSettings()
        {
            try
            {
                _userSettings = _backupService.GetLoadSettings(null);
                // TODO: refactor
                JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                try
                {
                    excel.ConnectFile();
                }
                catch (Exception)
                {
                    ;
                }

                // Если отчёт не отправлен в филиал. 
                if (DateTime.Now.Day > 15 && DateTime.Now.Day < 20) // TODO: вынести в настройки
                {
                    JwBookExcel.ArchiveReports archive = new JwBookExcel.ArchiveReports(excel);
                    var checkDate = JwBookExcel.ConvertDateTimeToStringArray(DateTime.Now);
                    if (archive.ContainsMonthReport(checkDate[0], checkDate[1]) == false)
                    {
                        MyMessageBox.Show("Не забудьте отправить отчёт в филиал!", "Уведомление");
                        _mainWindow.AddNotification(_mainWindow.CreateNotification("Уведомление", "Пожалуйста заполните и отправьте отчёт в филиал!"));
                    }
                }
                _mainWindow.AddNotification(_mainWindow.CreateNotification("Cистемное уведомление", "Настройки успешно загружены."));
                return _userSettings;
            }
            catch (InvalidOperationException ex) // Битый файл или неправильно сгенерированна розметка.
            {
                _waitWindow.Close();
                MyMessageBox.Show("Не удаеться правильно прочитать файл XML. Неккоректроне содержимое.", "Ошибка");
                _mainWindow.AddNotification(_mainWindow.CreateNotification("Cистемное уведомление", "Не удаеться правильно прочитать файл XML. Неккоректроне содержимое."));
            }
            catch (IndexOutOfRangeException ex) // Не нашли файл с настройками. Но директория уже создана.
            {
                var root = System.IO.Path.GetPathRoot(Environment.CurrentDirectory);
                _waitWindow.Close();
                MyMessageBox.Show($@"Не обнаружен файл настроект (XML) в директории: {root}MisistryReports\Settings", "Ошибка");
                _mainWindow.AddNotification(_mainWindow.CreateNotification("Cистемное уведомление", "Чтобы начать пользоваться программой - пожалуйста настройте. После того, как укажете все необходимые настройки - не забудьте сохранить. Если у вас уже есть файл настроек - пожалуйста загрузите его."));
                _mainWindow.AddNotification(_mainWindow.CreateNotification("Cистемное уведомление", $@"Не обнаружен файл XML в директории: {root}MisistryReports\Settings"));
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                // Create Dirictory -- Пользователь первый раз использует программу.
                _directoryService.CreateProgramDirectory();  // Создали необходимые системные директории.
                _directoryService.CreateSystemFile(); // Создаём копируем файлы, которые должны быть поумолчанию.
                
                _waitWindow.Close();
                _mainWindow.AddNotification(_mainWindow.CreateNotification("Системное уведомления", "Добро пожаловать! Ministry Reports это программа, которая поможет вам составлять и отправлять годовой отчёт в филиал. В это программе собраны все необходимые функции для удобного формирования годового отчета. Также вы можете отслеживать активность возвещателя, либо любого другого служителя собрания. Мы надеемся, что эта программа поможет вам формировать годовой отчет правильно и удобно. Чтобы начать пользоваться программой - настройте её."));
            }
            catch (NotSupportedException ex)
            {
                _waitWindow.Close();
                MyMessageBox.Show("JWBook: " + ex.Message, "Ошибка");
                _mainWindow.AddNotification(_mainWindow.CreateNotification("Cистемное уведомление", "JWBook: " + ex.Message));
            }

            return null;
        }
    }
}
