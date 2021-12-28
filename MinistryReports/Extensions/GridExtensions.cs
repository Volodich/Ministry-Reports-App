using System.Windows;
using System.Windows.Controls;

namespace MinistryReports.Extensions
{
    public static class GridExtensions
    {
        public static void Hidden(this Grid grid)
        {
            grid.Visibility = Visibility.Hidden;
            grid.IsEnabled = false;
        }

        public static void Visible(this Grid grid)
        {
            grid.Visibility = Visibility.Visible;
            grid.IsEnabled = true;
        }
    }
}
