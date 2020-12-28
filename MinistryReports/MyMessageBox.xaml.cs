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

namespace MinistryReports
{
    /// <summary>
    /// Логика взаимодействия для MyMessageBox.xaml
    /// </summary>
    public partial class MyMessageBox : Window
    {
        public MyMessageBox()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        public static void Show(string message, string title)
        {
            MyMessageBox messageBox = new MyMessageBox();
            messageBox.Title = title;
            messageBox.Message.Text = message;
            messageBox.ShowDialog();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            //TODO: сообщее продублировать в область уведомлений.
            await Task.Run(() => App.Current.Dispatcher.Invoke(() => this.Close()));
        }
    }
}
