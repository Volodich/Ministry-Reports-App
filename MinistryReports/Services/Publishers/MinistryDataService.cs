using MinistryReports.ViewModels;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace MinistryReports.Services.Publishers
{
    public interface IMinistryDataService
    {
        object[,] GetWorksheetData();
        int[] GetRangePublisher(string name);
        bool IsPublisherContainsInTable(string name);
        int GetColumnNumberInWorksheet(string month, string year);
        int GetStartYearIndex(string name, string year);
        void AddPublisher(string name, bool pioner = false, bool pastor = false, bool ministryAssistant = false);

        int GetMonthReports(string month, TypePublisher typePublisher,
            TypeGetPublisherResponceReport typeGetPublisherResponceReport, string year = "current");

        IEnumerable<string> PublsherWithoutReport(string month, string year);
        int CountActivePublishers();
        IEnumerable<List<object>> GetYearReportsPublisher(string name, string year);
        void UpdateDataPublisher(object[,] sendData, string name, string year);
    }

    public class MinistryDataService : JwBookExcel, IMinistryDataService
    {
        private ExcelPackage _package;
        private readonly UserSettings _settings;
        private readonly IPublisherDataManager _publisherDataManager;

        private int EndColumnData { get; set; }
        private int EndRowData { get; set; }
        public int StartMinistryYear { get; set; } // Служебный год, с которого начинаются данные в таблице

        public override ExcelWorksheet Worksheet { get; set; }
        public override string PathToWorkBook { get; }
        public override string NameTable { get; }

        public MinistryDataService(UserSettings settings)
        {
            _settings = settings;
            PathToWorkBook = _settings.JWBookSettings.JWBookPath;
            NameTable = ApplicationConfig.JwExcelBook.TableName;

            _publisherDataManager = new PublisherDataManager(settings.S21Settings);
            
            ConnectFile();

            var data = GetWorksheetData();
            EndColumnData = data.GetLength(1);
            EndRowData = data.GetLength(0);
            StartMinistryYear = GetStartYear(EndColumnData);
        }

        public override void ConnectFile()
        {
            FileInfo fileInfo = new FileInfo(PathToWorkBook);

            if (fileInfo.Exists)
            {
                _package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = _package.Workbook.Worksheets[NameTable];

                if (worksheet != null)
                {
                    Worksheet = worksheet;
                }
                else
                {
                    throw new NotSupportedException($"Таблица не найдена. Пожалуйста в Excel файле удостоверьтесь, что есть таблица с названием: {NameTable}");
                }
            }
        }

        public object[,] GetWorksheetData() => Worksheet.Cells.Value as object[,];
        
        private int GetStartYear(int endColumnData)
        {
            var columnsSymbs = GetColumnSymbolAsStringArray(endColumnData);
            int startIndex = 2; // 0 индекс - Имена, втоорой - описание отчёта (Публ., часы, видео и тп.)
            int monthCount = 12;
            int nextIndex = 1;

            var str = Worksheet.Cells[columnsSymbs[startIndex + monthCount + nextIndex] + "1"].Text;
            if (Int32.TryParse(str.Split('-').First(), out int year) == false)
            {
                if (year == default)
                    throw new Exception($"Не удалось провильно расспознать год. Обратите внимание на ячейку: {columnsSymbs[startIndex + monthCount + nextIndex]}1. Ячейка должна хранить значение служебного года. Например: 2018-2019.");
            }
            return year;
        }

        public int[] GetRangePublisher(string name)
        {
            int[] rangePublisher = null;

            var data = Worksheet.Cells.Value as object[,];
            for (int i = 1; i < EndRowData; i += 6)
            {
                if (data[i, 0].ToString() == name)
                {
                    rangePublisher = new int[2];
                    rangePublisher[0] = i;
                    rangePublisher[1] = i + 6;
                    break;
                }
            }
            if (rangePublisher == null)
            {
                throw new FileNotFoundException($"Возвещатель {name} не найден в таблице. Возможно указано неправильное имя или возвещателя нужно добавить.");
            }
            return rangePublisher;
        }

        public bool IsPublisherContainsInTable(string name)
        {
            var data = Worksheet.Cells.Value as object[,];
            for (int i = 1; i < EndRowData; i += 6)
            {
                if (data[i, 0].ToString() == name)
                {
                    return true;
                }
            }
            return false;
        }

        public int GetColumnNumberInWorksheet(string month, string year)
        {
            int intConvertMonth = 0;
            int start = 2; // месяцы начинаються с третьего ("D") столбца. 0-1-2 - системные столбцы.
            int yearCount = 12; // количество месяцев в году
            int sumColumn = 1; // строка подсчёта за год

            switch (month.ToLower())
            {
                case "январь":
                    intConvertMonth = 5 - 13;
                    break;
                case "февраль":
                    intConvertMonth = 6 - 13;
                    break;
                case "март":
                    intConvertMonth = 7 - 13;
                    break;
                case "апрель":
                    intConvertMonth = 8 - 13;
                    break;
                case "май":
                    intConvertMonth = 9 - 13;
                    break;
                case "июнь":
                    intConvertMonth = 10 - 13;
                    break;
                case "июль":
                    intConvertMonth = 11 - 13;
                    break;
                case "август":
                    intConvertMonth = 12 - 13;
                    break;
                case "сентябрь":
                    intConvertMonth = 1;
                    break;
                case "октябрь":
                    intConvertMonth = 2;
                    break;
                case "ноябрь":
                    intConvertMonth = 3;
                    break;
                case "декабрь":
                    intConvertMonth = 4;
                    break;
                default:
                    break;
            }

            if (Int32.TryParse(year, out int resultYear) == true)
            {
                return start + ((resultYear - StartMinistryYear) * (yearCount + sumColumn)) + intConvertMonth;
            }
            return -1;
        }

        public int GetStartYearIndex(string name, string year)
        {
            var data = Worksheet.Cells.Value as object[,];

            return GetColumnNumberInWorksheet("Сентябрь", year);

        }

        // TODO: refact: add publisher model
        public void AddPublisher(string name, bool pioner = false, bool pastor = false, bool ministryAssistant = false)
        {
            if (IsPublisherContainsInTable(name) == true)
            {
                throw new Exception($"Возвещатель {name} уже есть в таблице.");
            }

            var data = Worksheet.Cells.Value as object[,];
            
            List<string> names = new List<string>();
            for (int i = 1; i < EndRowData; i += 6)
            {
                names.Add(data[i, 0].ToString());
            }
            names.Add(name);
            names.Sort();
            var startIndex = names.IndexOf(name) * 6 + 2;
            var endIndex = startIndex + 6 + 1;


            // Данные которые будем заносить
            using (_package)
            {
                var ws = _package.Workbook.Worksheets[NameTable];
                ws.InsertRow(startIndex, 6);

                ws.Cells["A" + startIndex].Value = name;
                if (pioner == true)
                {
                    ws.Cells["A" + (startIndex + 2)].Value = "ПИОНЕР";
                    ws.Cells["A" + (startIndex + 2)].Style.Font.Bold = true;
                }
                if (pastor == true)
                {
                    ws.Cells["A" + (startIndex + 3)].Value = "СТАРЕЙШИНА";
                    ws.Cells["A" + (startIndex + 3)].Style.Font.Bold = true;
                }
                if (ministryAssistant == true)
                {
                    ws.Cells["A" + (startIndex + 3)].Value = "СЛУЖЕБНЫЙ ПОМОЩ.";
                    ws.Cells["A" + (startIndex + 3)].Style.Font.Bold = true;
                }
                ws.Cells["B" + startIndex].Value = "Публикации";
                ws.Cells["B" + (startIndex + 1)].Value = "Видео";
                ws.Cells["B" + (startIndex + 2)].Value = "Часы";
                ws.Cells["B" + (startIndex + 3)].Value = "Повт. Посещ.";
                ws.Cells["B" + (startIndex + 4)].Value = "Библ. Изуч.";
                ws.Cells["B" + (startIndex + 5)].Value = "Примечание";

                var arrColumnSym = GetColumnSymbolAsStringArray(EndColumnData);
                var lastColumnSym = arrColumnSym.Last();
                int monthCount = 12;

                ws.Cells[lastColumnSym + startIndex].Formula = $"SUM({arrColumnSym[arrColumnSym.Length - monthCount]}{startIndex}:{arrColumnSym[arrColumnSym.Length - 2]}{startIndex})";
                ws.Cells[lastColumnSym + (startIndex + 1)].Formula = $"SUM({arrColumnSym[arrColumnSym.Length - monthCount]}{(startIndex + 1)}:{arrColumnSym[arrColumnSym.Length - 2]}{(startIndex + 1)})";
                ws.Cells[lastColumnSym + (startIndex + 2)].Formula = $"SUM({arrColumnSym[arrColumnSym.Length - monthCount]}{(startIndex + 2)}:{arrColumnSym[arrColumnSym.Length - 2]}{(startIndex + 2)})";
                ws.Cells[lastColumnSym + (startIndex + 3)].Formula = $"SUM({arrColumnSym[arrColumnSym.Length - monthCount]}{(startIndex + 3)}:{arrColumnSym[arrColumnSym.Length - 2]}{(startIndex + 3)})";
                ws.Cells[lastColumnSym + (startIndex + 4)].Formula = $"SUM({arrColumnSym[arrColumnSym.Length - monthCount]}{(startIndex + 4)}:{arrColumnSym[arrColumnSym.Length - 2]}{(startIndex + 4)})";
                ws.Cells[lastColumnSym + (startIndex + 5)].Formula = $"SUM({arrColumnSym[arrColumnSym.Length - monthCount]}{(startIndex + 5)}:{arrColumnSym[arrColumnSym.Length - 2]}{(startIndex + 5)})";

                _package.Save();
            }

        }

        public int GetMonthReports(string month, TypePublisher typePublisher, TypeGetPublisherResponceReport typeGetPublisherResponceReport, string year = "current")
        {
            if (year == "current")
            { year = DateTime.Now.Year.ToString(); }

            int currentMonth = GetColumnNumberInWorksheet(month, year);
            if (currentMonth <= 0)
            { throw new NotSupportedException("Неправильно указана дата! Проверьте еще раз. Возможно данных за такой месяц/год еще нет?"); }

            int countPioner = 0;
            int countAPioner = 0;
            int countPublisher = 0;

            int countReportsPioner = 0;
            int countReportsAPioner = 0;
            int countReportsPublisher = 0;

            int loopStepPublisher = 6;
            TypePublisher tp = typePublisher;
            TypeGetPublisherResponceReport tgr;
            tgr = typeGetPublisherResponceReport != TypeGetPublisherResponceReport.count ? typeGetPublisherResponceReport : TypeGetPublisherResponceReport.hour;

            var data = Worksheet.Cells.Value as object[,];

            for (int i = 1; i < data.GetLength(0); i += 6)
            {
                if (data.GetLength(0) - 4 == i)
                {
                    break;
                }
                if (data[i, currentMonth] != null && data[i, currentMonth].ToString() != "")
                {
                    if (data[i + (loopStepPublisher - 1), currentMonth] != null && Regex.IsMatch(data[i + (loopStepPublisher - 1), currentMonth].ToString(), "\\bPP\\b"))
                    {
                        if (Int32.TryParse(data[i + (int)tgr, currentMonth].ToString(), out int result))
                        {
                            countAPioner += result;
                            countReportsAPioner++;
                        }
                        else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {data[i, 0].ToString()} за {month} {year}");
                    }
                    else if ((string)data[i + 2, 0] == "ПИОНЕР")
                    {
                        if (Int32.TryParse(data[i + (int)tgr, currentMonth].ToString(), out int result))
                        {
                            countPioner += result;
                            countReportsPioner++;
                        }
                        else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {data[i, 0].ToString()} за {month} {year}");
                    }
                    else
                    {
                        if (Int32.TryParse(data[i + (int)tgr, currentMonth].ToString(), out int result))
                        {
                            countPublisher += result;
                            countReportsPublisher++;
                        }
                        else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {data[i, 0].ToString()} за {month} {year}");
                    }
                }
                if (loopStepPublisher + i >= data.GetLength(0))
                {
                    break;
                }
            }
            switch (typeGetPublisherResponceReport)
            {
                case TypeGetPublisherResponceReport.count:
                    {
                        switch (typePublisher)
                        {
                            case TypePublisher.Publisher:
                                return countReportsPublisher;
                            case TypePublisher.AuxiliaryPioneer:
                                return countReportsAPioner;
                            case TypePublisher.Pioneer:
                                return countReportsPioner;
                            case TypePublisher.InactivePublisher:
                                return -1;
                            // TODO: return values
                            default:
                                return -1;
                        }
                    }
                default:
                    {
                        switch (typePublisher)
                        {
                            case TypePublisher.Publisher:
                                return countPublisher;
                            case TypePublisher.AuxiliaryPioneer:
                                return countAPioner;
                            case TypePublisher.Pioneer:
                                return countPioner;
                            case TypePublisher.InactivePublisher:
                                return -1;
                            // TODO: return values
                            default:
                                return -1;
                        }
                    }
            }
        }

        public IEnumerable<string> PublsherWithoutReport(string month, string year)
        {
            List<string> publishers = new List<string>();
            var data = Worksheet.Cells.Value as object[,];

            int currentMonth = GetColumnNumberInWorksheet(month, year);

            for (int i = 1; i < data.GetLength(0); i += 6)
            {
                if (i >= data.GetLength(0) - 4)
                {
                    break;
                }
                if (data[i + 2, currentMonth] == null || data[i + 2, currentMonth].ToString() == "")
                {
                    publishers.Add(data[i, 0].ToString());
                }
            }
            return publishers;
        }

        public int CountActivePublishers()
        {
            var data = Worksheet.Cells.Value as object[,];
            int count = default;
            for (int i = 3; i < data.GetLength(0); i += 6)
            {
                if (data[i, 0] != null && data[i, 0].ToString() == "NA")
                {
                    continue;
                }
                count++;
            }
            return count;
        }

        public IEnumerable<List<object>> GetYearReportsPublisher(string name, string year)
        {
            var data = Worksheet.Cells.Value as object[,];

            var rowIndex = GetRangePublisher(name);
            var columnIndex = GetColumnNumberInWorksheet("сентябрь", (Int32.Parse(year) - 1).ToString());

            List<object> publication = new List<object>();
            List<object> videos = new List<object>();
            List<object> hours = new List<object>();
            List<object> returnVisits = new List<object>();
            List<object> biblStudy = new List<object>();
            List<object> notice = new List<object>();

            for (int j = 0; j < 12; j++) // 12 - month count
            {
                if (data[rowIndex[0], columnIndex] == null) data[rowIndex[0], columnIndex] = String.Empty;
                publication.Add(data[rowIndex[0], columnIndex].ToString());
                if (data[rowIndex[0] + 1, columnIndex] == null) data[rowIndex[0] + 1, columnIndex] = String.Empty;
                videos.Add(data[rowIndex[0] + 1, columnIndex].ToString());
                if (data[rowIndex[0] + 2, columnIndex] == null) data[rowIndex[0] + 2, columnIndex] = String.Empty;
                hours.Add(data[rowIndex[0] + 2, columnIndex].ToString());
                if (data[rowIndex[0] + 3, columnIndex] == null) data[rowIndex[0] + 3, columnIndex] = String.Empty;
                returnVisits.Add(data[rowIndex[0] + 3, columnIndex].ToString());
                if (data[rowIndex[0] + 4, columnIndex] == null) data[rowIndex[0] + 4, columnIndex] = String.Empty;
                biblStudy.Add(data[rowIndex[0] + 4, columnIndex].ToString());
                if (data[rowIndex[0] + 5, columnIndex] == null) data[rowIndex[0] + 5, columnIndex] = String.Empty;
                notice.Add(data[rowIndex[0] + 5, columnIndex].ToString());
                columnIndex++;
            }

            return new List<List<object>>() { publication, videos, hours, returnVisits, biblStudy, notice };
        }

        public void UpdateDataPublisher(object[,] sendData, string name, string year)
        {
            var data = Worksheet.Cells.Value as object[,];

            int startIndexColumn = GetColumnNumberInWorksheet("сентябрь", (Int32.Parse(year) - 1).ToString());
            int startIndexRow = GetRangePublisher(name).First() + 1;

            using (_package)
            {
                var ws = _package.Workbook.Worksheets[NameTable];
                List<string> asas = new List<string>();
                int tempstartIColumn = startIndexColumn;
                var SymbolColumns = GetColumnSymbolAsStringArray(EndColumnData);
                for (int i = 0; i < sendData.GetLength(0); i++)
                {
                    for (int j = 0; j < sendData.GetLength(1); j++)
                    {
                        ws.Cells[$"{SymbolColumns[startIndexColumn]}{startIndexRow}"].Value = sendData[i, j];
                        startIndexColumn++;
                    }
                    startIndexColumn = tempstartIColumn; // Reset
                    startIndexRow++;
                }
                _package.Save();
            }
        }
    }

    /// <summary>
    /// Тип возвещателя
    /// </summary>
    public enum TypePublisher
    {
        Publisher, // возвещатель
        AuxiliaryPioneer, // подсобный пионер
        Pioneer, // пионер
        InactivePublisher, // неактивный возвещатель
        All // все возвещатели, не зависимо от назначения
    }

    /// <summary>
    /// Тип возвращаемых данных из отчётов (часы, публикации, повторные посещения)
    /// </summary>
    public enum TypeGetPublisherResponceReport
    {
        publications = 0,
        video,
        hour,
        returnReport,
        biblStudy,
        count
    }
}
