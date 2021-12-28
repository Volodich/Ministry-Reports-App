using System;
using System.Net;
using System.Net.Http;
using System.Windows.Documents;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using System.Runtime.CompilerServices;

namespace MinistryReports.Controllers
{
    public class FastWorkInstrumentsWPF
    {
        

        public class DataGrid
        {
            public static DataGridColumn CreateHeaderCollumn(string header)
            {
                DataGridColumn column = new DataGridTextColumn();
                column.Header = header;
                column.Visibility = Visibility.Visible;
                return column;
            }
        }

        public class RichTextBox
        {
            public static FlowDocument AddText(string text, FlowDocument document = null)
            {
                Paragraph paragraph = new Paragraph();
                paragraph.Inlines.Add(text);

                if (document != null)
                { 
                    document.Blocks.Add(paragraph);
                    return document;
                }
                FlowDocument flowDocument = new FlowDocument(paragraph);
                return flowDocument;
            }
        }

        public class WebRequest
        {
            //TODO Async Method
            public async static Task<bool> CheckInternetAsync()
            {
                return await Task<bool>.Factory.StartNew(CheckInternet); 
            }

            public static bool CheckInternet()
            {
                    try
                    {
                        using (var w = new WebClient())
                        {
                            w.DownloadStringAsync(new Uri("http://www.microsoft.com/"));
                        }
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                    return true;
            }
        }
    }
}
