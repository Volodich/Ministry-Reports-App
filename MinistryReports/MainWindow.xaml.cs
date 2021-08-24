using ConsoleApp4.JWBook;
using Microsoft.Win32;
using MinistryReports.Controllers;
using MinistryReports.Models;
using MinistryReports.Models.JWBook;
using MinistryReports.Models.S_21;
using MinistryReports.Serialization;
using MinistryReports.ViewModels;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO; // работа с файловой системой
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using MinistryReports.Extensions;
using MinistryReports.Services;

namespace MinistryReports
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private readonly IS21Manager _s21Manager;
        private readonly IS21Servise _s21Servise;
        private readonly IUserService _userService;

        internal UserSettings _userSettings;
        internal object dataPublisher;

        internal ObservableCollection<NoActivityPublisher> deletePublisher;
        internal ObservableCollection<JWMonthReport> meetreports;

        public MainWindow()
        {
            InitializeComponent();

            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            _userService = new UserService();
            _userSettings = _userService.GetUserSettings();
            _s21Servise = new S21Service();
            _s21Manager = new S21Manager(_userSettings.S21Settings);

            deletePublisher = new ObservableCollection<NoActivityPublisher>();
        }

        private async void MainWindowLoaded(object sender, EventArgs e)
        {
            if (sender is MainWindow mainWindow)
            {
                ProgressWindow waitWindow = new ProgressWindow
                {
                    ProgressBar =
                    {
                        IsIndeterminate = true, Orientation = System.Windows.Controls.Orientation.Horizontal
                    },
                    LabelInformation = {Content = String.Empty},
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this
                };

                waitWindow.Show(); // <--- Запускаем окно.

                Initialize startup = new Initialize(mainWindow, waitWindow);

                var uSettings =  startup.LoadAppSettings();
                if (uSettings != null)
                {
                    _userSettings = uSettings;
                    // если настройки загружены - пользоваться программой можно.
                    MonthReportWindow.IsEnabled = true;
                    S21Window.IsEnabled = true;
                    PublishersWindow.IsEnabled = true;
                    // Пока не реализовано нормально.
                    ArchiveMinistryWindow.IsEnabled = false;
                    ArchiveMinistryWindow.Visibility = Visibility.Hidden;

                    NoActivityWindow.IsEnabled = false;
                    NoActivityWindow.Visibility = Visibility.Hidden;
                }
                waitWindow.Close();
            }
            else
            {
                MyMessageBox.Show("Неверная загрузка работы программы! Приложение автоматически перезапустится!", "Ошибка");
                this.Close();
            }
        }

        private void ComboboxSelected(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            if (comboBox != null)
            {
                comboBox.Text = comboBox.SelectedItem.ToString();
            }
        }

        private void TextBoxMouseDoubleClickHandler(object sender, MouseButtonEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            textBox.Text = String.Empty;
            textBox.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
        }

        private int MonthConvertToInt(string month)
        {
            switch(month.ToLower())
            {
                case "январь":
                    return 1;
                case "февраль":
                    return 2;
                case "март":
                    return 3;
                case "апрель":
                    return 4;
                case "май":
                    return 5;
                case "июнь":
                    return 6;
                case "июль":
                    return 7;
                case "август":
                    return 8;
                case "сентябрь":
                    return 9;
                case "октябрь":
                    return 10;
                case "ноябрь":
                    return 11;
                case "декабрь":
                    return 12;
                default:
                    throw new FormatException("Не удалось расспознать месяц.");

            }
        }

        #region HomePage
        private void HamburgerMenuItemHomePage(object sender, MouseButtonEventArgs e)
        {
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(MonthReport);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(S21);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(PublisherInfo);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(Archive);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(SettingWindow);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(NoActivity);

            FastWorkInstrumentsWPF.HamburgerMenu.EnabledGridWindow(HomeWindow);
            MenuNameLabel.Content = "Главная";
        }

        #endregion

        #region MonhtReport Page
        private void HamburgerMenuItemMonthReportPage(object sender, RoutedEventArgs e)
        {
            // Visible & Enabled
            {
                FastWorkInstrumentsWPF.HamburgerMenu.EnabledGridWindow(MonthReport);

                MenuNameLabel.Content = "Месячный отчёт";

                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindowAsync(HomeWindow);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindowAsync(S21);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindowAsync(PublisherInfo);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindowAsync(Archive);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindowAsync(NoActivity);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(SettingWindow);

            }
            try
            {
                JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                excel.ConnectFile();

                JwBookExcel.DataPublisher dataPublisher = new JwBookExcel.DataPublisher(excel);

                int startYear = dataPublisher.StartMinistryYear;
                int currentYear = DateTime.Now.Year + 1;

                string[] years = new string[(currentYear - startYear) + 1]; // Всего массив хранить столько значений.
                for (int i = 0; i < years.Length; i++)
                { years[i] = currentYear.ToString(); currentYear--; }

                string[] months = new string[] { "Январь", "Февраль", "Март", "Апрель", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };

                List<string> yearList = new List<string>(); yearList.AddRange(years);
                List<string> monthList = new List<string>(); monthList.AddRange(months);

                ComboBoxYears.ItemsSource = yearList;
                ComboBoxMonth.ItemsSource = monthList;

                GetMonthReportButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                this.Dispatcher.Invoke(() => MyMessageBox.Show(ex.Message, "Ошибка"));
            }
        }

        private async void GetMonthReportButtonClick(object sender, RoutedEventArgs e)
        {
            SaveMonthReportButton.IsEnabled = false;

            string month = ComboBoxMonth.Text;
            string year = ComboBoxYears.Text;
            // Система покажет несдавших отчёт в том случае, если это текущий месяц а не прошлый. 
             
            bool currentFlag = false;
            if (DateTime.Now.Year.ToString() == year)
            {
                if (DateTime.Now.Month - MonthConvertToInt(month) == 1 || DateTime.Now.Month - MonthConvertToInt(month) == -11)
                {
                    currentFlag = true;
                }
            }

            ProgressWindow waitWindow = new ProgressWindow();
            waitWindow.ProgressBar.IsIndeterminate = true;
            waitWindow.ProgressBar.Orientation = System.Windows.Controls.Orientation.Horizontal;
            waitWindow.LabelInformation.Content = String.Empty;
            waitWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            waitWindow.Owner = this;


            await Task.Run(() => 
            { 
                if (month != String.Empty && year != String.Empty)
                {
                    this.Dispatcher.Invoke(()=> waitWindow.Show()); // <--- Запускаем окно.
                    try
                    {
                        JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                        JwBookExcel.DataPublisher dataPublisher = new JwBookExcel.DataPublisher(excel);
                        
                        ObservableCollection<JWMonthReport> monthReports = new ObservableCollection<JWMonthReport>();

                        monthReports.Add(new JWMonthReport() // Publisher
                        {
                            Type = "Возвещатель",
                            CountReports = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Publisher, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.count, year),
                            Publications = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Publisher, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.publications, year),
                            Videos = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Publisher, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.video, year),
                            Hours = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Publisher, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.hour, year),
                            ReturnVisits = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Publisher, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.returnReport, year),
                            BibleStudy = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Publisher, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.biblStudy, year),
                        });

                        monthReports.Add(new JWMonthReport() // APioner
                        {
                            Type = "Подсобный пионер",
                            CountReports = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.AuxiliaryPioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.count, year),
                            Publications = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.AuxiliaryPioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.publications, year),
                            Videos = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.AuxiliaryPioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.video, year),
                            Hours = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.AuxiliaryPioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.hour, year),
                            ReturnVisits = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.AuxiliaryPioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.returnReport, year),
                            BibleStudy = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.AuxiliaryPioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.biblStudy, year),
                        });

                        monthReports.Add(new JWMonthReport() // Pioner
                        {
                            Type = "Пионер",
                            CountReports = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Pioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.count, year),
                            Publications = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Pioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.publications, year),
                            Videos = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Pioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.video, year),
                            Hours = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Pioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.hour, year),
                            ReturnVisits = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Pioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.returnReport, year),
                            BibleStudy = dataPublisher.GetMonthReports(month, JwBookExcel.DataPublisher.TypePublisher.Pioneer, JwBookExcel.DataPublisher.TypeGetPublisherResponceReport.biblStudy, year),
                        });

                        monthReports.Add(new JWMonthReport() // Sum
                        {
                            Type = "Итого: ",
                            CountReports = monthReports[0].CountReports + monthReports[1].CountReports + monthReports[2].CountReports,
                            Publications = monthReports[0].Publications + monthReports[1].Publications + monthReports[2].Publications,
                            Videos = monthReports[0].Videos + monthReports[1].Videos + monthReports[2].Videos,
                            Hours = monthReports[0].Hours + monthReports[1].Hours + monthReports[2].Hours, 
                            ReturnVisits = monthReports[0].ReturnVisits + monthReports[1].ReturnVisits + monthReports[2].ReturnVisits,
                            BibleStudy = monthReports[0].BibleStudy + monthReports[1].BibleStudy + monthReports[2].BibleStudy
                        });

                        this.Dispatcher.Invoke(() => waitWindow.Close());

                        meetreports = monthReports; // записываем значение в поле класса.
                        if (meetreports != null)
                        {
                            this.Dispatcher.Invoke(() => MonthReportMeetDataGrid.ItemsSource = meetreports);
                        }

                        if (currentFlag == true) // true
                        {
                           
                            var publishersWithoutReport = dataPublisher.PublsherWithoutReport(month, year); // возвещатели, которые не сдели отчёт.
                            string arrPublisherWihtoutReport = String.Empty;
                            string longArrPublisherWihtoutReport = "Отчёт не здали: " + "\n";
                            for (int i = 0; i < publishersWithoutReport.Count; i++)
                            {
                                longArrPublisherWihtoutReport += (i + 1) + " " + publishersWithoutReport[i] + "\n";
                            }
                            if (publishersWithoutReport.Count < 5)
                            {
                                arrPublisherWihtoutReport = longArrPublisherWihtoutReport;
                            }
                            else
                            {
                                arrPublisherWihtoutReport = "Слишком много возвещателей не здало отчёт. Чтобы увидеть весь список зайдите в Главная и посмотрите список.";
                            }
                            this.Dispatcher.Invoke(() =>
                            { 
                                MyMessageBox.Show("Отчёт не здали: " + "\n" + arrPublisherWihtoutReport, "Месячный отчёт");
                                this.AddNotification(this.CreateNotification("Месячный отчёт", longArrPublisherWihtoutReport));
                                MyMessageBox.Show($"Количество активных возвещателей на {month} {year} год -- {dataPublisher.CountActivePublishers()}.", "Месячный отчёт");
                                });
    }
                        this.Dispatcher.Invoke(() => this.AddNotification(this.CreateNotification("Месячный отчёт", $"Количество активных возвещателей на {month} {year} год -- {dataPublisher.CountActivePublishers()}.")));
                        this.Dispatcher.Invoke(() => SaveMonthReportButton.IsEnabled = true);
                    }
                    catch (System.ArgumentOutOfRangeException ex) // Если произошла ошибка чтения данных с таблицы.  Графа суммирования данных отчёта пустая.
                    {
                        waitWindow.Close();
                        MyMessageBox.Show(ex.Message, "Ошибка");
                        this.AddNotification(this.CreateNotification("Ошибка", ex.Message));
                    }
                    catch (NotSupportedException ex) // Если в ячейках возвещателей не цифры а буквы. 
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            waitWindow.Close();
                            MyMessageBox.Show(ex.Message, "Ошибка");
                            this.AddNotification(this.CreateNotification("Ошибка", ex.Message));
                        });
                    }
                    catch (System.NullReferenceException)
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            waitWindow.Close();
                            MyMessageBox.Show("Ошибка при работе с таблицей. Пожалуйста проверьте правильнность данных в таблице и повторите попытку", "Ошибка");
                            this.AddNotification(this.CreateNotification("Ошибка", "Ошибка при работе с таблицей. Пожалуйста проверьте правильнность данных в таблице и повторите попытку."));
                        });
                    }
                    catch(System.IndexOutOfRangeException)
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            waitWindow.Close();
                            MyMessageBox.Show("Ошибка при работе с таблицей. Пожалуйста проверьте правильнность данных. Возможно данных за этот год еще нет в таблице?", "Ошибка");
                            this.AddNotification(this.CreateNotification("Ошибка", "Ошибка при работе с таблицей. Пожалуйста проверьте правильнность данных. Возможно данных за этот год еще нет в таблице?"));
                        });
                    }
                }
                else MyMessageBox.Show("Ошибка", "Укажите месяц и год!");
            });
        }

        private async void SaveMonthReportButtonClick(object sender, RoutedEventArgs e)
        {
            JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
            JwBookExcel.ArchiveReports archive = new JwBookExcel.ArchiveReports(excel);

            WarningWindow warning = new WarningWindow();
            warning.Owner = this;
            warning.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            var containtReport = await Task.Run(() => this.Dispatcher.Invoke(() => archive.ContainsMonthReport(ComboBoxMonth.Text, ComboBoxYears.Text)));
            if (containtReport == true) // Проверка есть ли данный отчёт уже в таблице.
            {
                this.Dispatcher.Invoke(() =>
                {
                    warning.TextBlockMessage.Text = "Отчёт за данный месяц уже есть в таблице. Вы хотите перезаписать данные?";
                    warning.WarningButton_OK.Click += UploadMonthReportWarningWindowBtnClick;
                    warning.ShowDialog();
                });
            }
            else
            {
                var data = meetreports;
                var sendData = new object[3, 6];
                for (int i = 0; i < data.Count-1; i++) // Ну учитываем Sum
                {
                    sendData[i, 0] = meetreports[i].CountReports;
                    sendData[i, 1] = meetreports[i].Publications;
                    sendData[i, 2] = meetreports[i].Videos;
                    sendData[i, 3] = meetreports[i].Hours;
                    sendData[i, 4] = meetreports[i].ReturnVisits;
                    sendData[i, 5] = meetreports[i].BibleStudy;
                }
                try
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        archive.CreateArchive(sendData, ComboBoxMonth.Text, ComboBoxYears.Text);
                        MyMessageBox.Show("Данные сохранены успешно!", "Месячный отчёт");
                        this.AddNotification(this.CreateNotification("Месячный отчёт", $"Данные за {ComboBoxMonth.Text} {ComboBoxYears.Text} сохранены успешно!"));
                    });
                }
                catch (System.InvalidOperationException ex) 
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        MyMessageBox.Show("Не удалось сохранить данные. Скорее всего у Вас открыта таблица JWBook. Закройте таблицу и повторите попытку!", "Ошибка!");
                        this.AddNotification(this.CreateNotification("Месячный отчёт", $"Не удалось сохранить данные.Скорее всего у Вас открыта таблица JWBook.Закройте таблицу и повторите попытку!"));
                    });
                }
                catch(Exception ex)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        MyMessageBox.Show($"Не удалось сохранить данные. Попробуйте позже! Причина: {ex.Message}", "Ошибка!");
                        this.AddNotification(this.CreateNotification("Месячный отчёт", $"Не удалось сохранить данные. Попробуйте позже! Причина: {ex.Message}"));
                    });
                }
            }
        }

        private async void UploadMonthReportWarningWindowBtnClick(object sender, RoutedEventArgs e)
        {
            JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
            JwBookExcel.ArchiveReports archive = new JwBookExcel.ArchiveReports(excel);

            var data = meetreports;
            var sendData = new object[3, 6];
            await Task.Run(() => 
            { 
                for (int i = 0; i < data.Count - 1; i++) // Ну учитываем Sum
                {
                    sendData[i, 0] = meetreports[i].CountReports;
                    sendData[i, 1] = meetreports[i].Publications;
                    sendData[i, 2] = meetreports[i].Videos;
                    sendData[i, 3] = meetreports[i].Hours;
                    sendData[i, 4] = meetreports[i].ReturnVisits;
                    sendData[i, 5] = meetreports[i].BibleStudy;
                }
                try
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        archive.UpdateArchive(sendData, ComboBoxMonth.Text, ComboBoxYears.Text);
                        MyMessageBox.Show("Данные сохранены успешно!", "Месячный отчёт");
                        this.AddNotification(this.CreateNotification("Месячный отчёт", $"Данные за {ComboBoxMonth.Text} {ComboBoxYears.Text} сохранены успешно!"));
                    });
                }
                catch (Exception)
                {
                    this.Dispatcher.Invoke(() => MyMessageBox.Show("Не удалось сохранить данные. Попробуйте позже!", "Ошибка!"));
                }
            });
        }

        private void MonthReportMeetDataGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            meetreports.Last().Publications = meetreports[0].Publications + meetreports[1].Publications + meetreports[2].Publications;
            meetreports.Last().Videos = meetreports[0].Videos + meetreports[1].Videos + meetreports[2].Videos;
            meetreports.Last().Hours = meetreports[0].Hours + meetreports[1].Hours + meetreports[2].Hours;
            meetreports.Last().ReturnVisits = meetreports[0].ReturnVisits + meetreports[1].ReturnVisits + meetreports[2].ReturnVisits;
            meetreports.Last().BibleStudy = meetreports[0].BibleStudy + meetreports[1].BibleStudy + meetreports[2].BibleStudy;
            meetreports.Last().CountReports = meetreports[0].CountReports + meetreports[1].CountReports + meetreports[2].CountReports;
            // Обновление данных  в таблице.
            MonthReportMeetDataGrid.ItemsSource = null;
            MonthReportMeetDataGrid.ItemsSource = meetreports;


            MyMessageBox.Show("Данные изменены!", "Успешно");
        }
        #endregion

        #region S-21 Page
        private void HamburgerMenuItemS21Page(object sender, MouseButtonEventArgs e)
        {
            // Visible & Enabled
            {
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(HomeWindow);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(MonthReport);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(PublisherInfo);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(Archive);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(NoActivity);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(SettingWindow);

                FastWorkInstrumentsWPF.HamburgerMenu.EnabledGridWindow(S21);
            }

            MenuNameLabel.Content = "Возвещатели";
            try
            {
                JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                JwBookExcel.DataPublisher dpExcel = new JwBookExcel.DataPublisher(excel);
                int startYear = dpExcel.StartMinistryYear;
                int currentYear = DateTime.Now.Year + 1;

                string[] years = new string[(currentYear - startYear) + 1]; // Всего массив хранить столько значений.
                for (int i = 0; i < years.Length; i++)
                { years[i] = currentYear.ToString(); currentYear--; }

                List<string> yearList = new List<string>(); yearList.AddRange(years);

                ComboBoxYearOld.ItemsSource = yearList;
                ComboBoxYearNow.ItemsSource = yearList;

                TextBoxPuthToFolderUnlaoding.Text = _userSettings.S21Settings.PuthToFolderUnlaoding;
            }
            catch (Exception ex)
            {
                this.Dispatcher.Invoke(() => MessageBox.Show(ex.Message, "Ошибка"));
            }
        }

        private void bthSearchFolder_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog browserDialog = new System.Windows.Forms.FolderBrowserDialog();
            browserDialog.Description = "Выберите папку для хранения pdf файлов с данными возвещателей собрания.";
            browserDialog.ShowDialog();
            TextBoxPuthToFolderUnlaoding.Text = browserDialog.SelectedPath;
        }

        private async void ButtonShowPublisher_Click(object sender, RoutedEventArgs e) // Показать всех возввещателей, карточки которым создадут.
        {
            try
            {
                if (ComboBoxYearOld.Text == String.Empty || ComboBoxYearNow.Text == String.Empty)
                    throw new NotSupportedException("Укажите правильно текущий служебный год и предыдущий.");
                if ((Int32.Parse(ComboBoxYearNow.Text) - Int32.Parse(ComboBoxYearOld.Text)) != 1)
                    throw new NotSupportedException("Для удачной операции, разница между текущим служебным годом и прошлым должна составлять один год! Проверьте правильно ли вы указали года!");

                var pd = await ExcelDBController.GetDataPublisherAsync(this._userSettings.S21Settings) as List<PublishersRange>;
                ObservableCollection<PublishersRange> publishersCollection = new ObservableCollection<PublishersRange>();
                foreach (var publisherData in pd)
                {
                    publishersCollection.Add(publisherData);
                }
                dataPublisher = pd;
                DataGridDataPublisher.Items.Clear();
                DataGridDataPublisher.ItemsSource = publishersCollection;

                ConvertToPDFButton.IsEnabled = true;
            }
            catch (NotSupportedException ex) // если неправильно указаны комбобоксы с годами.
            {
                MyMessageBox.Show(ex.Message, "Ошибка");
            }
            catch (FileLoadException ex) // если неправильно указан файл.
            {
                MyMessageBox.Show(ex.Message, "Ошибка");
                this.AddNotification(this.CreateNotification("Возвещатели", ex.Message));
            }
            catch (FileNotFoundException ex) // если неправильно указано местоположение файла.
            {
                MyMessageBox.Show(ex.Message, "Ошибка");
                this.AddNotification(this.CreateNotification("Возвещатели", ex.Message));
            }

        }

        private async void ConvertToPDFButton_Click(object sender, RoutedEventArgs e)
        {
            TextBoxPuthToFolderUnlaoding.BorderBrush = new SolidColorBrush(Color.FromRgb(0, 0, 0));

            var pd = dataPublisher as List<PublishersRange>; // pd - publisher data

            string firstYear = ComboBoxYearNow.Text;
            string secondYear = ComboBoxYearOld.Text;
            string publisherName;


            

            if (!_s21Servise.ExistTamplateFile)
            {
                MyMessageBox.Show("Не найден бланк S21! Проверьте путь к файлу указанный в настройках.", "Ошибка");
                this.AddNotification(this.CreateNotification("Возвещатели", "Не найден бланк S21! Проверьте путь к файлу указанный в настройках."));
                goto Exit;
            }

            var publisherInfoS21FieldFormat = _s21Manager.GetS21InfoPublisherFields(_userSettings.S21Settings).ToList();
            JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
            JwBookExcel.DataPublisher dataMinistryPublisher = new JwBookExcel.DataPublisher(excel);

            string puthFolder = String.Empty;
            
            if(TextBoxPuthToFolderUnlaoding.Text == String.Empty)
            {
                TextBoxPuthToFolderUnlaoding.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                MyMessageBox.Show("Укажите папку, в которой создадутся бланки.","Ошибка");
                goto Exit;
            }
            this.Dispatcher.Invoke(() =>
            {
                puthFolder = TextBoxPuthToFolderUnlaoding.Text;
            });

            ProgressWindow progressWindow = new ProgressWindow();
            progressWindow.LabelInformation.Content = String.Empty;
            progressWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            progressWindow.ProgressBar.Maximum = publisherInfoS21FieldFormat.Count;
            progressWindow.Owner = this;
            progressWindow.Show(); // <--- Запускаем окно.

            List<string> publisherNotFound = new List<string>();
            bool flagPublisherNotFound = false; // указывает - обнаружены ли ошибки. (В данном случае - те возвещатели - чби имена не совпадают.
            try
            {
                for (int i = 0; i < publisherInfoS21FieldFormat.Count(); i++)
                {
                    publisherName = _s21Manager.GetPublisherName(publisherInfoS21FieldFormat[i]);
                    try
                    {
                        var firstYearDataMinistry = dataMinistryPublisher.GetYearReportsPublisher(publisherName, firstYear);
                        var secondYearDataMinistry = dataMinistryPublisher.GetYearReportsPublisher(publisherName, secondYear);
                        var firstConvertDataMinistry = _s21Manager.CreateMinistryDataPublisherToStringFormat(firstYearDataMinistry, firstYear);
                        var secondConvertDataMinistry = _s21Manager.CreateMinistryDataPublisherToStringFormat(secondYearDataMinistry, secondYear);
                        _s21Manager.CreateDocument(_s21Manager.GetPublisherInfo(publisherInfoS21FieldFormat[i]),
                                                                firstConvertDataMinistry,
                                                                secondConvertDataMinistry,
                                                                puthFolder);
                        this.Dispatcher.Invoke(() => progressWindow.LabelInformation.Content = $"{i}/{publisherInfoS21FieldFormat.Count} -- {publisherName}");
                        this.Dispatcher.Invoke(() => progressWindow.ProgressBar.Value++);
                    }
                    catch (FileNotFoundException ex)
                    {
                        flagPublisherNotFound = true;
                        publisherNotFound.Add(publisherName);
                        continue;
                    }
                    catch(System.IO.IOException ex)
                    {
                        throw new Exception($"Не удалось найти файл Times New Roman.ttf. Пожалуйста переместите файл в папку {System.IO.Path.Combine(System.IO.Path.GetPathRoot(Environment.CurrentDirectory), "Ministry Reports", "Settings" )}.");
                    }
                }
                this.Dispatcher.Invoke(() => progressWindow.Close());
                MyMessageBox.Show($"Бланки S-21 успешно заполнены!", "Возвещатели");
                this.AddNotification(this.CreateNotification("Возвещатели", $"Бланки S-21 успешно заполнены! Всего создано {publisherInfoS21FieldFormat.Count - publisherNotFound.Count} бланков. Что бы просмотреть - перейдите - {puthFolder}"));
                if (flagPublisherNotFound) // true
                {
                    string message = "Возвещатели, для которых не удалось создать бланк: " + "\n";
                    for (int i = 0; i < publisherNotFound.Count; i++)
                    {
                        message += $"{i+1} - " + publisherNotFound[i] + "\n";
                    }
                    MyMessageBox.Show(message, "Возвещатели");
                    this.AddNotification(this.CreateNotification("Возвещатели", message));
                }
            }
            catch (NotSupportedException ex)
            {
                progressWindow.Close();
                MyMessageBox.Show(ex.Message, "Ошибка!");
                this.AddNotification(this.CreateNotification("Возвещатели", ex.Message));
            }
            catch (System.IO.DirectoryNotFoundException ex) // Если папка выгрузки пдф бланков не правильно указана
            {
                progressWindow.Close();
                TextBoxPuthToFolderUnlaoding.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                MyMessageBox.Show($"Провертие правильно ли вы указали папку создания pdf файлов.", "Ошибка!");
            }
            catch(Exception ex)
            {
                progressWindow.Close();
                MyMessageBox.Show(ex.Message, "Ошибка!");
                this.AddNotification(this.CreateNotification("Возвещатели", ex.Message));
            }
        Exit:;
        }
        #endregion

        #region PublisherInfo Page
        private async void HamburgerMenuItemPublisherInfoPage(object sender, MouseButtonEventArgs e)
        {
            // Visible & Enabled
            {
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(HomeWindow);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(S21);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(MonthReport);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(Archive);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(NoActivity);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(SettingWindow);

                MenuNameLabel.Content = "Информация о возвещателях";

                FastWorkInstrumentsWPF.HamburgerMenu.EnabledGridWindow(PublisherInfo);
            }

            ProgressWindow waitWindow = new ProgressWindow();
            waitWindow.ProgressBar.IsIndeterminate = true;
            waitWindow.ProgressBar.Orientation = System.Windows.Controls.Orientation.Horizontal;
            waitWindow.LabelInformation.Content = String.Empty;
            waitWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            waitWindow.Owner = this;
            waitWindow.Show(); // <--- Запускаем окно.

            await Task.Run(() =>
            {
                try
                {
                    JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                    JwBookExcel.DataPublisher dataPublisher = new JwBookExcel.DataPublisher(excel);

                    int startYear = dataPublisher.StartMinistryYear;
                    int currentYear = DateTime.Now.Year + 1;

                    string[] years = new string[(currentYear - startYear) + 1]; // Всего массив хранить столько значений.
                    for (int i = 0; i < years.Length; i++)
                    { years[i] = currentYear.ToString(); currentYear--; }

                    List<string> yearList = new List<string>(); yearList.AddRange(years);

                    this.Dispatcher.Invoke(() =>  ComboBoxYearsPublisherInfo.ItemsSource = yearList);

                    if (ExcelDBController.CheckConnect(_userSettings.S21Settings) == true)
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            waitWindow.Close();
                            SearchPublisherInfoButton.IsEnabled = true;
                            PublisherInfoAddPublisherButton.IsEnabled = true;
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            waitWindow.Close();
                            MyMessageBox.Show("Не удаёться подключиться к таблице Excel. Проверьте правильно ли вы настроили программу.", "Ошибка");
                        });
                    }
                }
                catch(Exception ex)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        waitWindow.Close();
                    MessageBox.Show("HamburgerMenuItemPublisherInfoPage Exception!");
                    });
                }
            });

        }

        private void DataGridDataMinistryPubliher_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            PublisherInfoSaveMinistryInformationButton.IsEnabled = true;
        }

        private async void SearchPublisherInfo_ClickHandler(object sender, RoutedEventArgs e)
        {
            if(TextBoxPublisherNamePublisherInfo.Text == String.Empty || TextBoxPublisherSurnamePublisherInfo.Text == String.Empty
                || TextBoxPublisherNamePublisherInfo.Text == "Введите имя" || TextBoxPublisherSurnamePublisherInfo.Text == "Введите фамилию")
            {
                MyMessageBox.Show("Заполните поле 'Имя' и 'Фамилия' для поиска возвещателя.", "Ошибка");
                goto exitMethod;
            }
            if(ComboBoxYearsPublisherInfo.Text == String.Empty || ComboBoxYearsPublisherInfo.Text == "Год")
            {
                MyMessageBox.Show("Укажите служебный год.", "Ошибка");
                goto exitMethod;
            }
            string namePublisher = $"{TextBoxPublisherSurnamePublisherInfo.Text} {TextBoxPublisherNamePublisherInfo.Text}";
            string year = ComboBoxYearsPublisherInfo.Text;
            await Task.Run(() =>
            {
                try
                {
                    JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                    JwBookExcel.DataPublisher dataMinistryPublisher = new JwBookExcel.DataPublisher(excel);
                    var yearMinistryData = dataMinistryPublisher.GetYearReportsPublisher(namePublisher, year);

                    ObservableCollection<JWMonthReport> publisherMD = new ObservableCollection<JWMonthReport>(); // Publisher Ministry Data
                    int monthCount = 12; // количество месяцев
                    string[] months = new string[] { "Сентябрь", "Октябрь", "Ноябрь", "Декабрь", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август" };
                    for (int i = 0; i < monthCount; i++)
                    {
                        if (yearMinistryData[0][i].ToString() == "")
                            yearMinistryData[0][i] = 0;
                        if (yearMinistryData[1][i].ToString() == "")
                            yearMinistryData[1][i] = 0;
                        if (yearMinistryData[2][i].ToString() == "")
                            yearMinistryData[2][i] = 0;
                        if (yearMinistryData[3][i].ToString() == "")
                            yearMinistryData[3][i] = 0;
                        if (yearMinistryData[4][i].ToString() == "")
                            yearMinistryData[4][i] = 0;
                        publisherMD.Add(new JWMonthReport()
                        {
                            Month = months[i],
                            Publications = Convert.ToInt32(yearMinistryData[0][i]),
                            Videos = Convert.ToInt32(yearMinistryData[1][i]),
                            Hours = Convert.ToInt32(yearMinistryData[2][i]),
                            ReturnVisits = Convert.ToInt32(yearMinistryData[3][i]),
                            BibleStudy = Convert.ToInt32(yearMinistryData[4][i]),
                            Notice = yearMinistryData[5][i].ToString()
                        });
                    }
                    dataPublisher = publisherMD;
                    this.Dispatcher.Invoke(() => DataGridDataMinistryPubliher.ItemsSource = publisherMD);
                }
                catch (NotSupportedException ex) // Если не удасться найти возвещателя в таблице.
                {
                    this.Dispatcher.Invoke(() =>  MyMessageBox.Show(ex.Message, "Ошибка"));
                }
                catch(System.IndexOutOfRangeException)
                {
                    this.Dispatcher.Invoke(() => MyMessageBox.Show("Кажется за этот год еще нет отчёта возвещателя.", "Ошибка"));
                }
            });
        exitMethod:;
        }

        private async void PublisherInfoSaveMinistryInformationButton_ClickHandler(object sender, RoutedEventArgs e)
        {
            var publisherMD = dataPublisher as ObservableCollection<JWMonthReport>; // Data Ministry Publisher
            
            string name = $"{TextBoxPublisherSurnamePublisherInfo.Text} {TextBoxPublisherNamePublisherInfo.Text}";
            string year = ComboBoxYearsPublisherInfo.Text;
            await Task.Run(() =>
            {
                JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                JwBookExcel.DataPublisher dpExcel = new JwBookExcel.DataPublisher(excel);

                object[,] updateData = new object[6, 12];
                for (int i = 0; i < 12; i++)
                {
                    updateData[0, i] = publisherMD[i].Publications;
                    updateData[1, i] = publisherMD[i].Videos;
                    updateData[2, i] = publisherMD[i].Publications;
                    updateData[3, i] = publisherMD[i].ReturnVisits;
                    updateData[4, i] = publisherMD[i].BibleStudy;
                    updateData[5, i] = publisherMD[i].Notice;
                }

                dpExcel.UpdateDataPublisher(updateData, name, year);
                this.Dispatcher.Invoke(() => MyMessageBox.Show("Данные обновлены!", "Успешно"));
            });
        }

        private void PublisherInfoAddPublisherButton_Click(object sender, RoutedEventArgs e)
        {
            PublisherWindow newPublisher = new PublisherWindow();
            newPublisher.Owner = this;
            newPublisher.ShowDialog();
        }

#endregion

        #region Archive Page
        private void HamburgerMenuItemArchivePage(object sender, MouseButtonEventArgs e)
        {
            // Visible & Enabled
            {
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(HomeWindow);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(S21);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(PublisherInfo);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(MonthReport);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(NoActivity);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(SettingWindow);

                // FastWorkInstrumentsWPF.HamburgerMenu.EnabledGridWindow(Archive);

                MenuNameLabel.Content = "Архив";
            }
            try
            {
                JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                JwBookExcel.DataPublisher dpExcel = new JwBookExcel.DataPublisher(excel);

                int startYear = dpExcel.StartMinistryYear;
                int currentYear = DateTime.Now.Year;

                string[] years = new string[(currentYear - startYear) + 1]; // Всего массив хранить столько значений.
                for (int i = 0; i < years.Length; i++)
                { years[i] = currentYear.ToString(); currentYear--; }

                string[] months = new string[] { "Январь", "Февраль", "Март", "Апрель", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };

                ComboBoxMeetArchiveYearFirst.ItemsSource = years;
                ComboBoxMeetArchiveYearSecond.ItemsSource = years;

                ComboBoxMeetArchiveMonthFirst.ItemsSource = months;
                ComboBoxMeetArchiveMonthSecond.ItemsSource = months;

                string[] typePublisher = new string[] { "Возвещатель", "Подсобный пионер", "Пионер", "Все" };

                ComboboxTypePublisher.ItemsSource = typePublisher;

                ButtonShowMeetReportsArchive.IsEnabled = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show("HamburgerMenuItemArchivePage Exception!");
            }
        }

        private async void ButtonClickShowMeetReportsArchive(object sender, RoutedEventArgs e)
        {
            try
            {
                JwBookExcel.DataPublisher dpExcel = new JwBookExcel.DataPublisher(new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth));
                
                if (ComboBoxMeetArchiveYearFirst.Text == String.Empty || ComboBoxMeetArchiveYearSecond.Text == String.Empty ||
                    ComboBoxMeetArchiveMonthFirst.Text == String.Empty || ComboBoxMeetArchiveMonthSecond.Text == String.Empty ||
                    ComboBoxMeetArchiveYearFirst.Text == "Год" || ComboBoxMeetArchiveYearSecond.Text == "Год" ||
                     ComboBoxMeetArchiveMonthFirst.Text == "Месяц" || ComboBoxMeetArchiveMonthSecond.Text == "Месяц")
                    throw new FormatException("Укажите правильно текущий служебный год и предыдущий.");
                if ((Int32.Parse(ComboBoxMeetArchiveYearFirst.Text) > Int32.Parse(ComboBoxMeetArchiveYearSecond.Text)))
                    throw new FormatException("Для удачной операции, начальный год должен быть больше, чем второй. Проверьте правильно ли вы указали года!");
                if (ComboboxTypePublisher.Text == String.Empty)
                    throw new FormatException("Укажите тип возвещателя, который необходимо вывести!");
                if ((Int32.Parse(ComboBoxMeetArchiveYearFirst.Text) == Int32.Parse(ComboBoxMeetArchiveYearSecond.Text)) && // Если года ровны
                    MonthConvertToInt(ComboBoxMeetArchiveMonthFirst.Text) > MonthConvertToInt(ComboBoxMeetArchiveMonthSecond.Text)) // Если первый месяц больше чем второй. Например Апрель (4) больше чем Март (3).
                    throw new FormatException("Проверьте правильно ли Вы указали месяцы. Возможно вы перепутали местами даты.");
                if( Convert.ToInt32(ComboBoxMeetArchiveYearFirst.Text) == dpExcel.StartMinistryYear&& MonthConvertToInt(ComboBoxMeetArchiveMonthFirst.Text) < 9) // 9 - Сентябрь - первый месяц начала нового служебного года. 
                {
                    throw new FormatException($"Данные за выбранный период не доступны. В Вашей Google таблице отчёты начинаються с сентября {dpExcel.StartMinistryYear} года .");
                }
                if (Convert.ToInt32(ComboBoxMeetArchiveYearSecond.Text) == dpExcel.StartMinistryYear && MonthConvertToInt(ComboBoxMeetArchiveMonthSecond.Text) < 9) // 9 - Сентябрь - первый месяц начала нового служебного года. 
                {
                    throw new FormatException($"Данные за выбранный период не доступны. В Вашей Google таблице отчёты начинаються с сентября {dpExcel.StartMinistryYear} года.");
                }

            }
            catch (FormatException ex)
            {
                MyMessageBox.Show($"{ex.Message}", "Ошибка");
                goto exitMethod;
            }
            
            string typePublisher = ComboboxTypePublisher.Text;
            string fYear = ComboBoxMeetArchiveYearFirst.Text;
            string sYear = ComboBoxMeetArchiveYearSecond.Text;
            string fMonth = ComboBoxMeetArchiveMonthFirst.Text;
            string sMonth = ComboBoxMeetArchiveMonthSecond.Text;

            ProgressWindow waitWindow = new ProgressWindow();
            waitWindow.ProgressBar.IsIndeterminate = true;
            waitWindow.ProgressBar.Orientation = System.Windows.Controls.Orientation.Horizontal;
            waitWindow.LabelInformation.Content = String.Empty;
            waitWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            waitWindow.Owner = this;
            waitWindow.Show(); // <--- Запускаем окно.

            try
            {
                JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
                JwBookExcel.ArchiveReports archive = new JwBookExcel.ArchiveReports(excel);

                var archiveReports = await Task.Run(() => archive.GetArchive(new string[] { fMonth, fYear, sMonth, sYear }, typePublisher));
                List<JWMonthReport> monthsReports = new List<JWMonthReport>();
                foreach (var ar in archiveReports)
                {
                    monthsReports.Add(new JWMonthReport() 
                    { 
                        Year = ar[0].ToString(),
                        Month = ar[1].ToString(),
                        Type = ar[2].ToString(),
                        CountReports = Int32.Parse(ar[3].ToString()),
                        Publications = Int32.Parse(ar[4].ToString()),
                        Videos = Int32.Parse(ar[5].ToString()),
                        Hours = Int32.Parse(ar[6].ToString()),
                        ReturnVisits = Int32.Parse(ar[7].ToString()),
                        BibleStudy = Int32.Parse(ar[8].ToString()),
                    });
                }
                this.Dispatcher.Invoke(() =>
                {
                    waitWindow.Close();
                    DataGridArchive.ItemsSource = monthsReports;
                });
            }
            catch (Exception ex)
            {
                this.Dispatcher.Invoke(() =>
                {
                    waitWindow.Close();
                    MyMessageBox.Show($"Непредвиденная ошибка при работе с данными Google таблици в листе {_userSettings.JWBookSettings.SheetNameArchiveReports}. Проверьте лист и его данные. '\n'Подробности: {ex.Message}. Action - ButtonClickShowMeetReportsArchive.", "Ошибка");
                    this.AddNotification(this.CreateNotification("Архив", "Не удалось подклоючиться к Google таблицам. Проверьте правильно ли Вы настроили приложение или подключение к интернету."));
                });
            }
        exitMethod:;
        }

        #endregion

        #region NoActivity Page
        private void HamburgerMenuItemNoActivityPage(object sender, MouseButtonEventArgs e)
        {
            {
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(HomeWindow);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(S21);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(PublisherInfo);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(Archive);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(MonthReport);
                FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(SettingWindow);

                // FastWorkInstrumentsWPF.HamburgerMenu.EnabledGridWindow(NoActivity);

                MenuNameLabel.Content = "Неактивные возвещатели";

                
            }
            try
            { 
                JwBookExcel.DataPublisher dpExcel = new JwBookExcel.DataPublisher(new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth));
                int startYear = dpExcel.StartMinistryYear;
                int currentYear = DateTime.Now.Year;

                string[] years = new string[(currentYear - startYear) + 1]; // Всего массив хранить столько значений.
                for (int i = 0; i < years.Length; i++)
                { years[i] = currentYear.ToString(); currentYear--; }

                YearNoActivityPablisherComboBox.ItemsSource = years;

                ShowNoActivityPublisherButton.IsEnabled = true;
            } 
            catch(Exception ex)
            { 
                MyMessageBox.Show(ex.Message, "Ошибка");
            }
        }

        private async void ShowNoActivityPublisherButton_Click(object sender, RoutedEventArgs e)
        {
            if(Int32.TryParse(YearNoActivityPablisherComboBox.Text, out int year) == false)
            {
                MyMessageBox.Show("Укажите правильно год!", "Ошибка");
                goto exitMethod;
            }
            JwBookExcel excel = new JwBookExcel(_userSettings.JWBookSettings.JWBookPuth);
            JwBookExcel.NoActivePublishers noActive = new JwBookExcel.NoActivePublishers(excel);

            var dataPublisher = await Task.Run(() => noActive.GetPublishers(year.ToString()));
            ObservableCollection<NoActivityPublisher> noActivityPublishers = new ObservableCollection<NoActivityPublisher>();

            foreach (var dp in dataPublisher)
            {
                noActivityPublishers.Add(new NoActivityPublisher() { Date = dp[0], Name = dp[1] });
            }

            NoActivityPublisherDataGrid.ItemsSource = noActivityPublishers;

        exitMethod:;
        }

        private void NoActivityPublisherDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            SaveChangeActivePublisherButton.IsEnabled = true;
        }

        private void NoActivityPublisherDataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var grid = (DataGrid)sender;
            if (Key.Delete == e.Key)
            {
                foreach (var row in grid.SelectedItems)
                {
                    deletePublisher.Add((NoActivityPublisher)row);
                }
                SaveChangeActivePublisherButton.IsEnabled = true;
            }

        }

        private async void SaveChangeActivePublisher_Click(object sender, RoutedEventArgs e)
        {
            var dataDelete = deletePublisher;
            deletePublisher = new ObservableCollection<NoActivityPublisher>(); // очистили список кого удалить
            // удаляем всех выбранных возвещателей.
            await Task.Run(() =>
            {
                try
                {
                    foreach (var deleteP in dataDelete)
                    {
                        //JWBookControllers.NoActivityPublisher.DeletePublisher(deleteP.Name, deleteP.Date);
                    }

                    var dataPublisher = NoActivityPublisherDataGrid.ItemsSource as ObservableCollection<NoActivityPublisher>;
                    ObservableCollection<string[]> dataPublisherConvert = new ObservableCollection<string[]>();
                    foreach (var dataP in dataPublisher)
                    {
                        dataPublisherConvert.Add(new string[] { dataP.Date, dataP.Name });
                    }
                    //JWBookControllers.NoActivityPublisher.EditPublisherData(dataPublisherConvert);
                }
                catch (SyntaxErrorException ex) // Sheet ID Exception
                {
                    this.Dispatcher.Invoke(() => {
                        MyMessageBox.Show(ex.Message, "Ошибка");
                        this.AddNotification(this.CreateNotification("Неактивные возвещатели", ex.Message));
                    });
                    goto exitMethod;
                }
                this.Dispatcher.Invoke(() =>
                {
                    SaveChangeActivePublisherButton.IsEnabled = false; // После того как сохранили - снова нужно внести изменение.
                    MyMessageBox.Show("Изменения успешно сохранены!", "Неактивные возвещатели");
                    this.AddNotification(this.CreateNotification("Неактивные возвещатели", "Изменения успешно сохранены"));
                });
            exitMethod:;
            });
        }

        #endregion

        #region Settings
        private void SettingsWindow_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(HomeWindow);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(S21);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(PublisherInfo);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(MonthReport);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(NoActivity);
            FastWorkInstrumentsWPF.HamburgerMenu.NoEnabledGridWindow(Archive);

            FastWorkInstrumentsWPF.HamburgerMenu.EnabledGridWindow(SettingWindow);

            MenuNameLabel.Content = "Настройки";

            InitializeTextBoxText(_userSettings);

        }
        

        // Заполнение полей во вкладке настройки доступной информацией.
        private void InitializeTextBoxText(UserSettings userSettings)
        {
            if (userSettings != null) // Если доступен файл настроек.
            {
                TextBoxUserName.Text = userSettings.UserName;
                TextBoxUserName.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));

                TextBoxJWBookExcelFile.Text = userSettings.JWBookSettings.JWBookPuth;
                TextBoxJWBookExcelFile.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                
                // S-21
                TextBoxPuthToUnloadingFolder.Text = userSettings.S21Settings.PuthToFolderUnlaoding;
                TextBoxPuthToUnloadingFolder.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));

                TextBoxPuthToPdfTamplateS21.Text = userSettings.S21Settings.PuthToTamplate;
                TextBoxPuthToPdfTamplateS21.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));

                TextBoxPuthToExcelFileDataPublisher.Text = userSettings.S21Settings.PuthToExcelDbFile;
                TextBoxPuthToExcelFileDataPublisher.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
            }
        }

        // Сработает, если пользователь изменит данные настроек.
        private void TextBoxTextChangedHandler(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (_userSettings != null)
            {
                ButtonReSavingSetting.IsEnabled = true;
                SaveSettings.IsEnabled = false;
            }
            textBox.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
        }

        private void ButtonReSavingSetting_Click(object sender, RoutedEventArgs e)
        {
            SaveSettings_Click(sender, e);
        }

        // Обработчик события - поиск файла или указание директории.
        private void ButtonClickOpenFileDialog(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.InitialDirectory = System.IO.Path.GetPathRoot(Environment.CurrentDirectory);

            TextBox textBox = new TextBox();
            Button btn = sender as Button;

            switch(btn.Name)
            {
                case "ButtonSearchSettingFile":
                    {
                        fileDialog.Filter = "XML files (*.xml) | *.xml;";
                        if (fileDialog.ShowDialog() == true)
                        {
                            string filePuth = fileDialog.FileName;
                            Backup backup = new Backup(_userSettings);
                            try
                            {
                                if (backup.LoadBackup(out _userSettings, filePuth)) // true
                                {
                                    // Даём доступ использованию программы.
                                    MonthReportWindow.IsEnabled = true;
                                    S21Window.IsEnabled = true;
                                    PublishersWindow.IsEnabled = true;
                                    ArchiveMinistryWindow.IsEnabled = true;
                                    NoActivityWindow.IsEnabled = true;

                                    InitializeTextBoxText(_userSettings);
                                }
                            }
                            catch (InvalidOperationException ex)
                            {

                                MyMessageBox.Show("Не удаеться правильно прочитать файл XML. Неккоректроне содержимое.", "Ошибка");
                                this.AddNotification(this.CreateNotification("Cистемное уведомление", "Не удаеться правильно прочитать файл XML. Неккоректроне содержимое."));
                            }
                            catch (Exception ex)
                            {
                                MyMessageBox.Show($"Неизвестная ошибка. Обратитесь к администратору.", "Ошибка");
                                this.AddNotification(this.CreateNotification("Неизвестная ошибка", $"Message: {ex.Message} InnerException: {ex.InnerException}. Window: SettingsWindow -> Button: ButtonSearchSettingFile"));
                            }
                        }
                    }
                    goto exitLabel;
                case "ButtonSearchJWBookFile":
                    textBox = TextBoxJWBookExcelFile;
                    fileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx;";
                    break;
                case "ButtonSearchFolderToUnload":
                    textBox = TextBoxPuthToUnloadingFolder;

                    System.Windows.Forms.FolderBrowserDialog browserDialog = new System.Windows.Forms.FolderBrowserDialog();
                    browserDialog.Description = "Выберите папку для хранения pdf файлов с данными возвещателей собрания.";
                    if(browserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        textBox.Text = browserDialog.SelectedPath;
                        textBox.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                    }
                    goto exitLabel;
                case "ButtonSearchPdfS21Blank":
                    textBox = TextBoxPuthToPdfTamplateS21;
                    fileDialog.Filter = "PDF files (*.pdf) | *.pdf;";
                    break;
                case "ButtonSearchExcelFile":
                    textBox = TextBoxPuthToExcelFileDataPublisher;
                    fileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx;";
                    break;
            }
            if (fileDialog.ShowDialog() == true)
            {
                textBox.Text = fileDialog.FileName;
                textBox.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
            }
        exitLabel:;
        }

        // Обработчик события - кнопка сохранить настройки.
        private void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            ProgressWindow waitWindow = new ProgressWindow();
            waitWindow.ProgressBar.IsIndeterminate = true;
            waitWindow.ProgressBar.Orientation = System.Windows.Controls.Orientation.Horizontal;
            waitWindow.LabelInformation.Content = String.Empty;
            waitWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            waitWindow.Owner = this;
            waitWindow.Show(); // <--- Запускаем окно.

            Task.Factory.StartNew(() => 
            {
                Thread.Sleep(1000);
                this.Dispatcher.Invoke(() =>
                {
                    if (TextBoxUserName.Text != null &&
                        // JWBook
                        TextBoxJWBookExcelFile.Text != "Путь к файлу" &&
                        // S-21
                        TextBoxPuthToUnloadingFolder.Text != "Путь к папке" &&
                        TextBoxPuthToPdfTamplateS21.Text != "Путь к месту хранения бланка" &&
                        TextBoxPuthToExcelFileDataPublisher.Text != "Путь к месту хранения файла excel")
                    {
                        bool flagFilePuth = true;
                        // проверка пути к файлам.
                        if (!TextBoxJWBookExcelFile.Text.Contains(@"\"))
                        {
                            waitWindow.Close();
                            MyMessageBox.Show("Ошибка - не удаёться найти файл по указанному пути", "Ошибка");
                            TextBoxJWBookExcelFile.Foreground = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                            flagFilePuth = false;
                        }
                        if (!TextBoxPuthToUnloadingFolder.Text.Contains(@"\"))
                        {
                            waitWindow.Close();
                            MyMessageBox.Show("Ошибка - не удаёться определить путь", "Ошибка");
                            TextBoxPuthToUnloadingFolder.Foreground = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                            flagFilePuth = false;
                        }
                        if (!TextBoxPuthToExcelFileDataPublisher.Text.Contains(@"\"))
                        {
                            waitWindow.Close();
                            MyMessageBox.Show("Ошибка - не удаёться найти файл по указанному пути.", "Ошибка");
                            TextBoxPuthToExcelFileDataPublisher.Foreground = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                            flagFilePuth = false;
                        }
                        if (!TextBoxPuthToPdfTamplateS21.Text.Contains(@"\"))
                        {
                            waitWindow.Close();
                            MyMessageBox.Show("Ошибка - не удаёться найти файл по указанному пути.", "Ошибка");
                            TextBoxPuthToPdfTamplateS21.Foreground = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                            flagFilePuth = false;
                        }

                        // если успешно заполнены поля.
                        if (flagFilePuth == true)
                        {
                            bool flagEditFile = true;
                            if (_userSettings == null) // Первый раз сохраняем настройки.
                            {
                                _userSettings = new UserSettings() { JWBookSettings = new Models.JWBookSettings(), S21Settings = new S21Settings() };
                                flagEditFile = false;
                            }
                            try
                            {
                                // заносим данные в экземпляр класса UserSettings
                                _userSettings.UserName = TextBoxUserName.Text;
                                //JwBook
                                _userSettings.JWBookSettings.JWBookPuth = TextBoxJWBookExcelFile.Text;
                                // S-21
                                _userSettings.S21Settings.PuthToFolderUnlaoding = TextBoxPuthToUnloadingFolder.Text;
                                _userSettings.S21Settings.PuthToTamplate = TextBoxPuthToPdfTamplateS21.Text;
                                _userSettings.S21Settings.PuthToExcelDbFile = TextBoxPuthToExcelFileDataPublisher.Text;

                                // Сохранение в файл.
                                Backup backup = new Backup(_userSettings);
                                backup.CreateBackup();

                                // Открываем доступ к использованию программы.
                                MonthReportWindow.IsEnabled = true;
                                S21Window.IsEnabled = true;
                                PublishersWindow.IsEnabled = true;
                                ArchiveMinistryWindow.IsEnabled = true;
                                NoActivityWindow.IsEnabled = true;

                                // Закрываем окно ожидания.
                                waitWindow.Close();

                                // Уведомляем пользователя об супешной операции.
                                MyMessageBox.Show($"Данные сохранены!", "Успешно");
                                if(flagEditFile) //
                                    this.AddNotification(this.CreateNotification("Настройки", "Изменения сохранены!"));
                                else
                                    this.AddNotification(this.CreateNotification("Настройки", $@"Файл с настройками успешно создан! Вы можете его найти по пути: {System.IO.Path.GetPathRoot(Environment.CurrentDirectory)}Ministry Reports\Settings"));

                            }
                            catch (Exception ex) // если произойдёт непредвиденная ошибка при заполнении полей пользователем.
                            {
                                waitWindow.Close();
                                MyMessageBox.Show($"Неизвестная ошибка. Обратитесь к администратору.", "Ошибка");
                                this.AddNotification(this.CreateNotification("Неизвестная ошибка", $"Message: {ex.Message} InnerException: {ex.InnerException}. Window: SettingsWindow -> Button: ButtonSaveSettingFile"));
                            }
                        }
                        else {
                            waitWindow.Close();
                            MyMessageBox.Show($"Сохранение не удалось. Попробуйте еще раз. Рекомендуем проверить правильно ли заполнены все поля.", "Ошибка");
                        }
                    }

                    else
                    {
                        waitWindow.Close();
                        MyMessageBox.Show("Заполните пожалуйста все поля", "Ошибка");
                    }
                });
            }); 
        }



        #endregion

       
    }
}
