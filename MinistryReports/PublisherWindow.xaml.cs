using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MinistryReports.Models;
using MinistryReports.Models.S_21;
using MinistryReports.Controllers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System.Globalization;
using ConsoleApp4.JWBook;

namespace MinistryReports
{
    /// <summary>
    /// Логика взаимодействия для AddPublisher.xaml
    /// </summary>
    public partial class PublisherWindow : Window
    {
        private Dictionary<string, string> Months { get; set; } // Месяца в комбобокс

        public PublisherWindow()
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            ComboboxInitialize();
        }

        void ComboboxInitialize()
        {
            int currentYear = DateTime.Now.Year;
            List<string> years = new List<string>();
            List<string> days = new List<string>();

            for (int i = 0; i < 100; i++) // 100 - максимальный возраст возвещателя
            {
                years.Add((currentYear--).ToString());
                if(i < 31) // 31 - количество дней в месяце
                {
                    days.Add(i.ToString());
                }
            }

            Months = new Dictionary<string, string>();
            Months.Add("Января", "01");
            Months.Add("Февраля", "02");
            Months.Add("Марта","03");
            Months.Add("Апреля","04");
            Months.Add("Майя","05");
            Months.Add("Июня","06");
            Months.Add("Июля","07");
            Months.Add("Августа","08");
            Months.Add("Сентября","09");
            Months.Add("Октября","10");
            Months.Add("Ноября","11");
            Months.Add("Декабря","12");

            DayComboBoxBaptism.ItemsSource = days;
            MonthComboBoxBaptism.ItemsSource = Months.Keys;
            YearComboBoxBaptism.ItemsSource = years;

            DayComboBoxBirth.ItemsSource = days;
            MonthComboBoxBirth.ItemsSource = Months.Keys;
            YearComboBoxBirth.ItemsSource = years;
        }

        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (TextboxName.Text == String.Empty ||
                TextboxSurname.Text == String.Empty ||
                TextboxAddress.Text == String.Empty ||                
                TextBoxMobile1.Text == String.Empty ||
                TextBoxMobile2.Text == String.Empty
                )
            { MyMessageBox.Show("Необходимо заполнить все поля!", "Ошибка!"); goto exitMethod; }
            if ((CheckBoxMen.IsChecked == false && CheckBoxWomen.IsChecked == false) ||
                    (CheckBoxMen.IsChecked == true && CheckBoxWomen.IsChecked == true))
            { MyMessageBox.Show("Выберите один вариант пола (либо мужчина либо женщина)", "Ошибка!"); goto exitMethod; }
            if (CheckBoxPastor.IsChecked == true && CheckBoxMinistryHelper.IsChecked == true)
            { MyMessageBox.Show("Брат может иметь талько одно назначение!", "Ошибка!"); goto exitMethod; }
            if (DayComboBoxBirth.Text == "День" || DayComboBoxBaptism.Text == "День")
            { MyMessageBox.Show("Проверьте правильно ли указан день рождения (крещения).", "Ошибка!"); goto exitMethod; }
            if (MonthComboBoxBirth.Text == "Месяц" || MonthComboBoxBaptism.Text == "Месяц")
            { MyMessageBox.Show("Проверьте правильно ли указан месяц рождения (крещения).", "Ошибка!"); goto exitMethod; }
            if (YearComboBoxBirth.Text == "Год" || YearComboBoxBaptism.Text == "Год")
            { MyMessageBox.Show("Проверьте правильно ли указан год рождения (крещения).", "Ошибка!"); goto exitMethod; }
            {
                if(Int32.TryParse(YearComboBoxBirth.Text, out int yearBirth) == false)
                {
                    MyMessageBox.Show("Укажите правильный год рождения!","Ошибка");
                    goto exitMethod;
                }
                if (Int32.TryParse(YearComboBoxBaptism.Text, out int yearBaptism) == false)
                {
                    MyMessageBox.Show("Укажите правильный год крещения!","Ошибка");
                    goto exitMethod;
                }
                if(yearBaptism < yearBirth) // Если год крещения больше года рождения - ошибка
                {
                    MyMessageBox.Show("Год рождения не может быть меньше, чем год крещения!", "Ошибка");
                    goto exitMethod;
                }
            }


            string birthPublisher = DayComboBoxBirth.Text + "." + Months[MonthComboBoxBirth.Text] + "." + YearComboBoxBirth.Text;
            string baptismPublisher = DayComboBoxBaptism.Text + "." + Months[MonthComboBoxBaptism.Text] + "." + YearComboBoxBaptism.Text;
            PublishersRange publisher = new PublishersRange()
            {
                Name = TextboxSurname.Text + " " + TextboxName.Text,
                Adress = TextboxAddress.Text,
                DateBirth = birthPublisher,
                BuptismDate = baptismPublisher,
                Mobile1 = TextBoxMobile1.Text,
                Mobile2 = TextBoxMobile2.Text,
                Gender = CheckBoxMen.IsChecked == true ? "М" : "Ж",
                Appointment = CheckBoxPastor.IsChecked == true ? "СТАР" : (CheckBoxMinistryHelper.IsChecked == true ? "СЛУЖ" : ""),
                Pioner = CheckBoxPioner.IsChecked == true ? "П" : String.Empty
            };

            ProgressWindow waitWindow = new ProgressWindow();
            waitWindow.ProgressBar.IsIndeterminate = true;
            waitWindow.ProgressBar.Orientation = System.Windows.Controls.Orientation.Horizontal;
            waitWindow.LabelInformation.Content = String.Empty;
            waitWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            waitWindow.Owner = this;
            waitWindow.Show();

            await Task.Run(() =>
            {
                MainWindow mainWindow = null;
                this.Dispatcher.Invoke(() => mainWindow = this.Owner as MainWindow);
                JwBookExcel excel = new JwBookExcel(mainWindow.userSettings.JWBookSettings.JWBookPuth);
                JwBookExcel.DataPublisher dpExcel = new JwBookExcel.DataPublisher(excel);
                if (dpExcel.IsPublisherContainsInTable(publisher.Name) == true)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        waitWindow.Close();
                        MyMessageBox.Show("Такой возвещатель уже есть в системе!", "Ошибка!");
                    });
                }
                else
                {
                    try
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            dpExcel.AddPublisher(publisher.Name, pioner: CheckBoxPioner.IsChecked.Value, pastor: CheckBoxPastor.IsChecked.Value, ministryAssistant: CheckBoxMinistryHelper.IsChecked.Value);
                            ExcelDBController.AddPublisher(publisher, mainWindow.userSettings.S21Settings);
                            waitWindow.Close();
                            MyMessageBox.Show($"Возвещатель \"{publisher.Name}\" успешно сохранён!", "Успешно!");
                            MainWindow window = this.Owner as MainWindow;
                            window.AddNotification(MainWindow.CreateNotification("Информация о возвещателях", $"Возвещатель \"{publisher.Name}\" успешно сохранён в Google таблицы и Excel файл."));
                        });
                    }
                    catch (InvalidOperationException)
                    {
                        waitWindow.Close();
                        MyMessageBox.Show($"Не удалось добавить возвещателя \"{publisher.Name}\". Проверьте настройки приложения. Возможно Вы не правильно настроили доступ к файлу excel. Или возможно у Вас открыт документ excel. Пожалуйста закройте его и повторите попытку.", "Ошибка!");
                        MainWindow window = this.Owner as MainWindow;
                        window.AddNotification(MainWindow.CreateNotification("Информация о возвещателях", $"Не удалось добавить возвещателя \"{publisher.Name}\". Проверьте настройки приложения. Возможно Вы не правильно настроили доступ к файлу excel. Или возможно у Вас открыт документ excel. Пожалуйста закройте его и повторите попытку."));
                    }
                }
            });
            this.Close();
        exitMethod:;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ComboboxSelected(object sender, RoutedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            if (comboBox != null)
            {
                comboBox.Text = comboBox.SelectedItem.ToString();
                comboBox.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
            }
        }

        private void TextboxTextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            textBox.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
        }

        private void TextboxDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            textBox.Text = String.Empty;
        }
    }
}
