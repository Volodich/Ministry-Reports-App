using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MinistryReports.Models;
using MinistryReports.ViewModels;
using MinistryReports.JWBook;
using Google.Apis.Sheets.v4.Data;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Runtime.InteropServices.WindowsRuntime;
using Org.BouncyCastle.Asn1.Mozilla;
using System.Collections.ObjectModel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using MinistryReports.Models.JWBook;
using Google;
using System.Runtime.CompilerServices;
using System.Net.Http;
using System.ComponentModel;
using System.CodeDom;

namespace MinistryReports.Controllers
{
    class JWBookControllers
    {
        public static bool IsConnected { get => true; } // TODO: нормально сделать.

        public static bool IsMonthReports { get => JWBook.JWBook.MeetReport.CheckMonthReport(DateTime.Now); }

        public static int StartYear { get => JWBook.JWBook.StartYear; }

        public static List<string> PublisherWithoutReport(string month, string year)
        {
             JWBook.JWBook.DataMinistry.GetPublisherReports.Month monthR = new JWBook.JWBook.DataMinistry.GetPublisherReports.Month();
            return monthR.PublisherWithoutReport(month, year);
        }

        public static List<JWMonthReport> MonthReportRequest(string month, string year)
        {
            List<JWMonthReport> monthReportMeet = new List<JWMonthReport>();
            try
            {
                JWBook.JWBook.DataMinistry.GetPublisherReports.Month monthReport = new JWBook.JWBook.DataMinistry.GetPublisherReports.Month();
                monthReportMeet.Add(new JWMonthReport
                {
                    Publications = monthReport.Publications(month, TypePublisher.Publisher, year),
                    Videos = monthReport.Videos(month, TypePublisher.Publisher, year),
                    Hours = monthReport.Hours(month, TypePublisher.Publisher, year),
                    ReturnVisits = monthReport.ReturnVisits(month, TypePublisher.Publisher, year),
                    BibleStudy = monthReport.BiblStudy(month, TypePublisher.Publisher, year),
                    CountReports = monthReport.CountReports(month, TypePublisher.Publisher, year),

                    Type = "Возвещатель"
                });
                monthReportMeet.Add(new JWMonthReport
                {
                    Publications = monthReport.Publications(month, TypePublisher.AuxiliaryPioneer, year),
                    Videos = monthReport.Videos(month, TypePublisher.AuxiliaryPioneer, year),
                    Hours = monthReport.Hours(month, TypePublisher.AuxiliaryPioneer, year),
                    ReturnVisits = monthReport.ReturnVisits(month, TypePublisher.AuxiliaryPioneer, year),
                    BibleStudy = monthReport.BiblStudy(month, TypePublisher.AuxiliaryPioneer, year),
                    CountReports = monthReport.CountReports(month, TypePublisher.AuxiliaryPioneer, year),

                    Type = "Подсобный пионер"
                });
                monthReportMeet.Add(new JWMonthReport
                {
                    Publications = monthReport.Publications(month, TypePublisher.Pioneer, year),
                    Videos = monthReport.Videos(month, TypePublisher.Pioneer, year),
                    Hours = monthReport.Hours(month, TypePublisher.Pioneer, year),
                    ReturnVisits = monthReport.ReturnVisits(month, TypePublisher.Pioneer, year),
                    BibleStudy = monthReport.BiblStudy(month, TypePublisher.Pioneer, year),
                    CountReports = monthReport.CountReports(month, TypePublisher.Pioneer, year),

                    Type = "Пионер"
                });
                JWMonthReport sum = new JWMonthReport()
                {
                    Publications = monthReportMeet[0].Publications + monthReportMeet[1].Publications + monthReportMeet[2].Publications,
                    Videos = monthReportMeet[0].Videos + monthReportMeet[1].Videos + monthReportMeet[2].Videos,
                    Hours = monthReportMeet[0].Hours + monthReportMeet[1].Hours + monthReportMeet[2].Hours,
                    ReturnVisits = monthReportMeet[0].ReturnVisits + monthReportMeet[1].ReturnVisits + monthReportMeet[2].ReturnVisits,
                    BibleStudy = monthReportMeet[0].BibleStudy + monthReportMeet[1].BibleStudy + monthReportMeet[2].BibleStudy,
                    CountReports = monthReportMeet[0].CountReports + monthReportMeet[1].CountReports + monthReportMeet[2].CountReports,
                    CountActivePublishers = monthReport.CountActivePublisher(month, year),
                    Type = "Итого: "
                };
                monthReportMeet.Add(sum);
            }
            catch (System.ArgumentOutOfRangeException ex)
            {
                //MessageBox.Show($"Ошибка! System.ArgumentOutOfRangeException JWBookControllers.MonthReportRequest. Подробности: {ex.Message}", "Ошибка");
                throw new System.ArgumentOutOfRangeException($"Ошибка при чтении данных из таблицы с отчётами возвещателй. Проверьте - заполнены ли поля суммы данных отчётов. (См. стр. ХХ). Или проверьте правильно ли Вы ввели дату. (См. стр. ХХ).");
            }
            catch (NotSupportedException ex) // если в ячейках возвещателей не цифры а буквы. 
            {
                //MessageBox.Show($"Ошибка! NotSupportedException JWBookControllers.MonthReportRequest. Подробности: {ex.Message}", "Ошибка");
                throw ex;
            }
            return monthReportMeet; 
        }

        public static void RefreshSpreadsheet(JWBookSettings settings)
        {
            JWBook.JWBook jwb = new JWBook.JWBook(settings);
        }

        public static Task RefreshSpreadsheetAsync(JWBookSettings settings)
        {
            return Task.Run(() => { RefreshSpreadsheet(settings); });
        }

        public async static Task<List<JWMonthReport>> MonthReportRequestAsync( string month, string year)
        {
            return await Task.Run(() => MonthReportRequest(month, year));
        }

        public static bool CheckContainsReport(string month, string year)
        {
            JWBook.JWBook.MeetReport meetReport = new JWBook.JWBook.MeetReport();
            return meetReport.ItContains(month, year);
        }

        public static async Task<bool> CheckContainsReportAsync(string month, string year)
        {
            return await Task.Run(() => CheckContainsReport( month, year));
        }

        public static List<IList<object>> CreateValueRangeToArchiveMeetReport(object meetReports, string month, string year)
        {
            var reports = meetReports as ObservableCollection<JWMonthReport>;
            if (reports == null)
            { return null; }

            List<object> oblist;
            List<IList<object>> valueRange = new List<IList<object>>();
            for (int i = 0; i < 3; i++)
            {
                oblist = new List<object>();
                oblist.Add(year);
                oblist.Add(month);
                oblist.Add(reports[i].Type);
                oblist.Add(reports[i].CountReports);
                oblist.Add(reports[i].Publications);
                oblist.Add(reports[i].Videos);
                oblist.Add(reports[i].Hours);
                oblist.Add(reports[i].ReturnVisits);
                oblist.Add(reports[i].BibleStudy);
                valueRange.Add(oblist);
            }
            return valueRange;
        }

        public static List<IList<object>> CreateValueRangeToMinistryReportsPublisher(object reports)
        {
            var ministryReports = reports as ObservableCollection<JWMonthReport>;
            if (ministryReports == null)
            {
                return null;
            }

            List<object> publications = new List<object>();
            List<object> videos = new List<object>();
            List<object> hours = new List<object>();
            List<object> returnVisits = new List<object>();
            List<object> bibleStudy = new List<object>();
            List<object> notice = new List<object>();

            List<IList<object>> valueRange = new List<IList<object>>();
            int iterationCount = 12;
            for (int i = 0; i < iterationCount-1; i++)
            {
                publications.Add(ministryReports[i].Publications);
                videos.Add(ministryReports[i].Videos);
                hours.Add(ministryReports[i].Hours);
                returnVisits.Add(ministryReports[i].ReturnVisits);
                bibleStudy.Add(ministryReports[i].BibleStudy);
                notice.Add(ministryReports[i].Notice);
            }

            valueRange.Add(publications);
            valueRange.Add(videos);
            valueRange.Add(hours);
            valueRange.Add(returnVisits);
            valueRange.Add(bibleStudy);
            valueRange.Add(notice);

            return valueRange;
        }

        public static List<IList<object>> CreateValueRangeToAddPublisher(string name, string hopeOther, bool pioner = false, string appointment = "")
        {
            List<object> column1 = new List<object>();
            string[] namep = name.Split(' ');
            if (namep.Length < 2)
                column1.Add($"{namep[0]}");
            else
            column1.Add($"{namep[0]} {namep?[1]}");
            column1.Add("Публикации");
            
            List<object> column2 = new List<object>();
            column2.Add("");
            column2.Add("Видео");

            List<object> column3 = new List<object>();
            if (pioner == true)
                column3.Add("ПИОНЕР");
            else column3.Add("");
            column3.Add("Часы");
            
            List<object> column4 = new List<object>();
            column4.Add(appointment);
            column4.Add("Повт. Посещ.");
            
            List<object> column5 = new List<object>();
            column5.Add("");
            column5.Add("Библ. Изуч.");

            List<object> column6 = new List<object>();
            column6.Add("");
            column6.Add("Примечание");

            List<IList<object>> valueRange = new List<IList<object>>();
            valueRange.Add(column1);
            valueRange.Add(column2);
            valueRange.Add(column3);
            valueRange.Add(column4);
            valueRange.Add(column5);
            valueRange.Add(column6);

            return valueRange;
        }

        public static bool UpdateMonthReportsMeet(object valueRange ,string month, string year)
        {
            var vRange = valueRange as List<IList<object>>;

            JWBook.JWBook.MeetReport meetReport = new JWBook.JWBook.MeetReport();
            meetReport.UpdateArchive(vRange, month, year);
            return true;
        }

        public static bool CreateMonthReportsMeet(object valueRange, string month, string year)
        {
            var vRange = valueRange as List<IList<object>>;

            JWBook.JWBook.MeetReport meetReport = new JWBook.JWBook.MeetReport();
            meetReport.CreateArchive(vRange, month, year);
            return true;
        }

        public static bool UpdateMinistryDataPublisher(object valueRange, string name, string year)
        {
            return JWBook.JWBook.DataMinistry.UpdatePublisherReports.UpdateYearData(valueRange, name, year);
        }

        public async static Task<bool> UpdateMinistryDataPublisherAsync(object valueRange, string name, string year)
        {
            return await Task.Run(() => UpdateMinistryDataPublisher(valueRange, name, year));
        }

        public static List<List<object>> GetDataMinistryPublisher(object dataMinistry,string name, string year)
        {
            ValueRange valueRange = dataMinistry as ValueRange;
            return JWBook.JWBook.DataMinistry.GetOfflineDataPublisher(valueRange, name, year); // dp = DataPublishers
        }

        public static object DataMinistryAll(JWBookSettings settings)
        {
            return JWBook.JWBook.DataMinistry.GetAllDataPublisher(settings.SpreadsheetId);
        }

        public static Task<object> DataMinistryAllAsync(JWBookSettings settings)
        {
            return Task.Run(() => DataMinistryAll(settings));
        }

        public static List<IList<object>> GetMeetReportsArchive(string[] selection, string typePublisher)
        {
            try
            {
                JWBook.JWBook.MeetReport meetReport = new JWBook.JWBook.MeetReport();
                return meetReport.GetArchive(selection, typePublisher);
            }
            catch(Google.GoogleApiException)
            {
                throw new NotSupportedException();
            }
            catch(AggregateException)
            {
                throw new NotSupportedException();
            }
        }

        public async static Task<List<JWMonthReport>> GetMeetReportsArchiveAsync(string[] selection, string typePublisher)
        {
            var reports = await Task.Run(() => GetMeetReportsArchive(selection, typePublisher));

            List <JWMonthReport> response = new List<JWMonthReport>();
            foreach (var report in reports)
            {
                response.Add(new JWMonthReport()
                {
                    Year = ((List<object>)report)[0].ToString(),
                    Month = ((List<object>)report)[1].ToString(),
                    Type = ((List<object>)report)[2].ToString(),
                    CountReports = Int32.Parse(((List<object>)report)[3].ToString()),
                    Publications = Int32.Parse(((List<object>)report)[4].ToString()),
                    Videos = Int32.Parse(((List<object>)report)[5].ToString()),
                    Hours = Int32.Parse(((List<object>)report)[6].ToString()),
                    ReturnVisits = Int32.Parse(((List<object>)report)[7].ToString()),
                    BibleStudy = Int32.Parse(((List<object>)report)[8].ToString())
                });
            }
            return response;
        }

        public static bool CheckPublisherContain(string name)
        {
            return JWBook.JWBook.DataMinistry.CheckPublisher(name);
        }

        public static Task<bool> CheckPublisherContainAsync(string name)
        {
            return Task.Run(() => CheckPublisherContain(name));
        }

        public static void AddPublisher(string name, string hopeOther = "д.о.", bool pioner = false, string appointment = "")
        {
            var allData = JWBook.JWBook.DataMinistry.GetAllDataPublisher();
            int countNewPublisher = 6; // новый возвещатель каждые 6 элементов
            
            List<string> publishers = new List<string>();

            for (int i = 0; i < allData.Count; i+=countNewPublisher)
            {
                publishers.Add(allData[i][0].ToString());
            }
            publishers.Add(name);
            publishers.Sort();
            int indexNewPublisher = publishers.IndexOf(name);
            var valueRange = CreateValueRangeToAddPublisher(name, hopeOther, pioner, appointment);
            JWBook.JWBook.DataMinistry.AddPblisher(valueRange, indexNewPublisher);
        }

        public static void SetNRPublisher(List<string> publishers, string month, string year)
        {
            JWBook.JWBook.DataMinistry.GetPublisherReports.Month.SetNRPublisher(publishers, month, year);
        }

        public class NoActivityPublisher
        {
            private static JWBook.JWBook.NoActivityPublisher instance;

            public static void ConnectToSheetList()
            {
                try
                {
                    instance = new JWBook.JWBook.NoActivityPublisher();
                }
                catch (Google.GoogleApiException)
                {
                    throw new System.Data.SyntaxErrorException("Ошибка подключения к Google таблице. Проверьте настройки приложения.");
                }
                catch(FormatException)
                {
                    throw new System.Data.SyntaxErrorException("Ошибка при обработке данных в листе нективных возвещателей. Проверьте правильность данных.");
                }
                
            }

            public static ObservableCollection<string[]> GetNoActivityPublisher(string year = "current")
            {
                try
                {
                    return instance.GetRequest(year);
                }
                catch (System.ArgumentOutOfRangeException ex)
                {
                    //MessageBox.Show($"Ошибка! System.ArgumentOutOfRangeException JWBookControllers.NoActivityPublisher.GetNoActivityPublisher. Подробности: {ex.Message}", "Ошибка");
                    throw new System.ArgumentOutOfRangeException($"Ошибка при чтении данных из таблицы с отчётами возвещателй. Проверьте - заполнены ли поля суммы данных отчётов. (См. стр. ХХ). Или проверьте правильно ли Вы ввели дату. (См. стр. ХХ).");
                }
                catch (NotSupportedException ex) // если в ячейках возвещателей не цифры а буквы. 
                {
                    //MessageBox.Show($"Ошибка! NotSupportedException JWBookControllers.NoActivityPublisher.GetNoActivityPublisher Подробности: {ex.Message}", "Ошибка");
                    throw ex;
                }
            }

            public static void EditPublisherData(ObservableCollection<string[]> updateData)
            {
                try
                {
                    instance.UpdateRequest(updateData);
                }
                catch (Google.GoogleApiException)
                {
                    throw new System.Data.SyntaxErrorException("Ошибка подключения к Google таблице. Проверьте настройки приложения.");
                }
            }

            public static void DeletePublisher(string name, string date)
            {
                try
                {
                    instance.DeleteRequest(name, date);
                }
                catch (Google.GoogleApiException)
                {
                    throw new System.Data.SyntaxErrorException("Ошибка подключения к Google таблице. Проверьте настройки приложения. Обратите внимание на SheetId листа неактивных возвещателей.");
                }
            }

            public static void CreatePublisher(string name, string date)
            {
                instance.CreateRequest(name, date);
            }

            public static bool CheckPublisher(string name, string date)
            {
                return instance.CheckPublisher(name, date);
            }

            public static void Refresh()
            {
                try
                {
                    instance.RefreshData();
                }
                catch(Google.GoogleApiException)
                {
                    throw new System.Data.SyntaxErrorException("Ошибка подключения к Google таблице. Проверьте настройки приложения.");
                }
                
            }
        }
    }

}
