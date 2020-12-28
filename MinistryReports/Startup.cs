using System;
using System.Threading.Tasks;
using System.Windows;
using MinistryReports.Models;
using MinistryReports.ViewModels;
using MinistryReports.Serialization;
using MinistryReports.Controllers;
using System.Threading;
using System.IO;
using System.Runtime.CompilerServices;
using System.Security;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using ConsoleApp4.JWBook;

namespace MinistryReports
{
    /// <summary>
    /// Все операции, которые должны произойти при запуске программы.
    /// </summary>
    class Startup
    {

        private MainWindow mainWindow;
        private Window waitWindow;
        public UserSettings userSettings;

        public Startup(MainWindow window, Window waitWindow)
        {
            this.mainWindow = window;
            this.waitWindow = waitWindow;
        }

        public async Task<UserSettings> StartAsync()
        {
            
            Backup backup = new Backup(userSettings);
            try // Обработка ошибок связанных с файловой системой и загрузки настроек.
            {
                if (!backup.LoadBackup(out userSettings))
                {
                    return null; // мы не можем загрузить настройки. Нет возможности и подключиться к таблицам
                }

                JwBookExcel excel = new JwBookExcel(userSettings.JWBookSettings.JWBookPuth);
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
                        mainWindow.AddNotification(MainWindow.CreateNotification("Уведомление", "Пожалуйста заполните и отправьте отчёт в филиал!"));
                    }
                }
                mainWindow.AddNotification(MainWindow.CreateNotification("Cистемное уведомление", "Настройки успешно загружены."));
                return userSettings;
            }
            catch (InvalidOperationException ex) // Битый файл или неправильно сгенерированна розметка.
            {
                waitWindow.Close();
                MyMessageBox.Show("Не удаеться правильно прочитать файл XML. Неккоректроне содержимое.", "Ошибка");
                mainWindow.AddNotification(MainWindow.CreateNotification("Cистемное уведомление", "Не удаеться правильно прочитать файл XML. Неккоректроне содержимое."));
            }
            catch (IndexOutOfRangeException ex) // Не нашли файл с настройками. Но директория уже создана.
            {
                var root = System.IO.Path.GetPathRoot(Environment.CurrentDirectory);
                waitWindow.Close();
                MyMessageBox.Show($@"Не обнаружен файл настроект (XML) в директории: {root}MisistryReports\Settings", "Ошибка");
                mainWindow.AddNotification(MainWindow.CreateNotification("Cистемное уведомление", "Чтобы начать пользоваться программой - пожалуйста настройте. После того, как укажете все необходимые настройки - не забудьте сохранить. Если у вас уже есть файл настроек - пожалуйста загрузите его."));
                mainWindow.AddNotification(MainWindow.CreateNotification("Cистемное уведомление", $@"Не обнаружен файл XML в директории: {root}MisistryReports\Settings"));
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                // Create Dirictory -- Пользователь первый раз использует программу.
                CreateProgramDirectory();  // Создали необходимые системные директории.
                waitWindow.Close();
                mainWindow.AddNotification(MainWindow.CreateNotification("Системное уведомления", "Добро пожаловать! Ministry Reports это программа, которая поможет вам составлять и отправлять годовой отчёт в филиал. В это программе собраны все необходимые функции для удобного формирования годового отчета. Также вы можете отслеживать активность возвещателя, либо любого другого служителя собрания. Мы надеемся, что эта программа поможет вам формировать годовой отчет правильно и удобно. Чтобы начать пользоваться программой - настройте её."));
            }
            catch(NotSupportedException ex)
            {
                waitWindow.Close();
                MyMessageBox.Show("JWBook: " + ex.Message, "Ошибка");
                mainWindow.AddNotification(MainWindow.CreateNotification("Cистемное уведомление", "JWBook: " + ex.Message));
            }
            //catch (Exception ex)
            //{
            //    waitWindow.Close();
            //    MyMessageBox.Show($"Непредвиденная ошибка. Подробности: {ex.Message}", "Ошибка");
            //}
            return null;
        }

        /// <summary>
        /// Метод который создаёт дериктории необходимые. Рекомендуется использовать если программу запустили впервые (или нет настроект полдьзователя)
        /// </summary>
        public void CreateProgramDirectory()
        {
            var root = Path.GetPathRoot(Environment.CurrentDirectory);
            string programFolder = "Ministry Reports";
            Directory.CreateDirectory(Path.Combine(root, programFolder));
            Directory.CreateDirectory(Path.Combine(root, programFolder, "Settings"));
            Directory.CreateDirectory(Path.Combine(root, programFolder, "Reports Publishers"));
        }
    }
}
