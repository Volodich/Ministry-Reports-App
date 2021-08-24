using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace MinistryReports.Extensions
{
    public static class NotoficationExtensions
    {
        public static FlowDocument CreateNotification(this Window window, string title, string body, double bodyTextIndent = 0)
        {
            FlowDocument notification = new FlowDocument();

            Run titleNotification = new Run(title);
            titleNotification.Foreground = new SolidColorBrush(Color.FromRgb(139, 139, 139));
            titleNotification.FontFamily = new System.Windows.Media.FontFamily("Roboto");
            Paragraph firstParagraph = new Paragraph();
            firstParagraph.Inlines.Add(titleNotification);
            firstParagraph.TextAlignment = TextAlignment.Left;

            Run bodyNotification = new Run(body);
            bodyNotification.Foreground = new SolidColorBrush(Color.FromRgb(44, 44, 44));
            bodyNotification.FontStyle = FontStyles.Normal;
            bodyNotification.FontSize = 16.0;
            bodyNotification.FontFamily = new System.Windows.Media.FontFamily("Roboto");
            Paragraph secondParagraph = new Paragraph();
            secondParagraph.Inlines.Add(bodyNotification);
            secondParagraph.TextAlignment = TextAlignment.Justify;
            secondParagraph.TextIndent = bodyTextIndent;

            Run bottomInformation = new Run(DateTime.Now.ToString("HH:mm dd.MM"));
            bottomInformation.Foreground = new SolidColorBrush(Color.FromRgb(139, 139, 139));
            Paragraph thirdParagraph = new Paragraph();
            thirdParagraph.Inlines.Add(bottomInformation);
            thirdParagraph.TextAlignment = TextAlignment.Right;

            notification.Blocks.Add(firstParagraph);
            notification.Blocks.Add(secondParagraph);
            notification.Blocks.Add(thirdParagraph);

            return notification;
        }

        internal static void AddNotification(this MainWindow window, FlowDocument notification)
        {
            RichTextBox richTextBox = new RichTextBox();
            richTextBox.Document = notification;
            richTextBox.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            richTextBox.BorderBrush = null;
            richTextBox.IsReadOnly = true;
            Thickness thickness = new Thickness();
            thickness.Bottom = 0;
            richTextBox.BorderThickness = thickness;
            richTextBox.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            Thickness paddingText = new Thickness() { Left = 16, Right = 10 };
            richTextBox.Padding = paddingText;
            window.NotificationStackPanel.Children.Add(richTextBox);

            RichTextBox separator = new RichTextBox();
            separator.Height = 10;
            separator.Background = null;
            separator.IsEnabled = false;
            window.NotificationStackPanel.Children.Add(separator);

        }
    }
}
