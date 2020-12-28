using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.ObjectModel;

namespace ConsoleApp4.JWBook
{
    internal class JwBookExcel
    {
        private string puth;
        public string PuthToWorkbook { get => puth; }
        public static string Alphabet { get => "A-B-C-D-E-F-G-H-I-J-K-L-M-N-O-P-Q-R-S-T-U-V-W-X-Y-Z"; }

        private ExcelWorksheet dataMinistryWorkSheet;
        private ExcelWorksheet archiveReportsWorkSheet;
        private ExcelWorksheet noActivePublisherWorkSheet;

        internal ExcelWorksheet DataMinistryWorkSheet { get => dataMinistryWorkSheet; }
        internal ExcelWorksheet ArchiveReportsWorkSheet { get => archiveReportsWorkSheet; }
        internal ExcelWorksheet NoActivePublisherWorkSheet { get => noActivePublisherWorkSheet; }

        private string NameTableDataMinistry { get => "Reports"; }
        private string NameTableArchiveReports { get => "Archive"; }
        private string NameTableNoActivePublisher { get => "NoActivePublisher"; }

        private ExcelPackage package;

        public JwBookExcel(string puth)
        {
            this.puth = puth;
            ConnectFile();
        }

        public void ConnectFile()
        {
            FileInfo fileInfo = new FileInfo(PuthToWorkbook);
            if (fileInfo.Exists == true)
            {
                package = new ExcelPackage(fileInfo);
                ExcelWorksheet datamonistryWB = package.Workbook.Worksheets[NameTableDataMinistry];
                if (datamonistryWB != null)
                {
                    this.dataMinistryWorkSheet = datamonistryWB;
                }
                else throw new NotSupportedException($"Таблица не найдена. Пожалуйста в Excel файле удостоверьтесь, что есть таблица с названием: {NameTableDataMinistry}");
                ExcelWorksheet archiveReportsWB = package.Workbook.Worksheets[NameTableArchiveReports];
                if (archiveReportsWB != null)
                {
                    this.archiveReportsWorkSheet = archiveReportsWB;
                }
                else throw new NotSupportedException($"Таблица не найдена. Пожалуйста в Excel файле удостоверьтесь, что есть таблица с названием: {NameTableArchiveReports}");
                ExcelWorksheet noActivePublisherWB = package.Workbook.Worksheets[NameTableNoActivePublisher];
                if (noActivePublisherWB != null)
                {
                    this.noActivePublisherWorkSheet = noActivePublisherWB;
                }
                else throw new NotSupportedException($"Таблица не найдена. Пожалуйста в Excel файле удостоверьтесь, что есть таблица с названием: {NameTableNoActivePublisher}");
            }
            //else throw new FileNotFoundException($"Не удалось найти файл по указанному пути: {PuthToWorkbook}. Пожалуйста проверьте наличие файла. Если файл открыт закройте его.");

        }

        public static string[] GetColumnSymbolAsStringArray(int count)
        {
            string[] arrSymbolColumn = new string[count];
            string[] alphabetSym = Alphabet.Split('-');
            int iteratorAlphabet = 0;
            int countRepeate = -1;
            string symRepeate = String.Empty;

            for (int i = 0; i < arrSymbolColumn.Length; i++)
            {
                if (iteratorAlphabet >= alphabetSym.Length)
                {
                    iteratorAlphabet = 0;
                    countRepeate++;
                    symRepeate = alphabetSym[countRepeate];
                }
                arrSymbolColumn[i] = symRepeate + alphabetSym[iteratorAlphabet];
                iteratorAlphabet++;
            }
            return arrSymbolColumn;
        }

        public static int ConvertMonthStringToInt32(string month)
        {
            switch (month)
            {
                case "Январь":
                    return 1;
                case "Февраль":
                    return 2;
                case "Март":
                    return 3;
                case "Апрель":
                    return 4;
                case "Май":
                    return 5;
                case "Июнь":
                    return 6;
                case "Июль":
                    return 7;
                case "Август":
                    return 8;
                case "Сентябрь":
                    return 9;
                case "Октябрь":
                    return 10;
                case "Ноябрь":
                    return 11;
                case "Декабрь":
                    return 12;
                default:
                    throw new Exception("Не удалось определить месяц!");
            }
        }

        public static string[] ConvertDateTimeToStringArray(DateTime date)
        {
            var month = date.Month.ToString(); //ConvertMonthStringToInt32(date.Month.ToString()).ToString();
            var year = date.Year.ToString();
            return new string[] { month, year };
        }

        public class DataPublisher
        {
            private int EndColumnData { get; set; }
            private int EndRowData { get; set; }
            public int StartMinistryYear { get; set; } // Служебный год, с которого начинаются данные в таблице.
            public static int StaticStartMinistryYear { get; set; } 
            private string[] SymbolColumns { get; set; }

            private ExcelWorksheet worksheet;
            public readonly string puthToFile;
            public readonly string nameTable;

            public DataPublisher(JwBookExcel excelData)
            {
                puthToFile = excelData.PuthToWorkbook;
                nameTable = excelData.NameTableDataMinistry;

                worksheet = excelData.dataMinistryWorkSheet;
                var data = worksheet.Cells.Value as object[,];
                EndColumnData = data.GetLength(1);
                EndRowData = data.GetLength(0);

                StartMinistryYear = GetStartYear(EndColumnData, excelData.DataMinistryWorkSheet);
                StaticStartMinistryYear = StartMinistryYear;
            }

            public int GetStartYear(int endColumnData, ExcelWorksheet data)
            {
                var columnsSymbs = GetColumnSymbolAsStringArray(endColumnData);
                int startIndex = 2; // 0 индекс - Имена, втоорой - описание отчёта (Публ., часы, видео и тп.)
                int monthCount = 12;
                int nextIndex = 1;

                var str = data.Cells[columnsSymbs[startIndex + monthCount + nextIndex] + "1"].Text;
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

                var data = worksheet.Cells.Value as object[,];
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
                var data = worksheet.Cells.Value as object[,];
                for (int i = 1; i < EndRowData; i += 6)
                {
                    if (data[i, 0].ToString() == name)
                    {
                        return true;
                    }
                }
                return false;
            }

            public static int ConvertMonthYearToIndexInArray(string month, string year)
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
                    return start + ((resultYear - StaticStartMinistryYear) * (yearCount + sumColumn)) + intConvertMonth;
                }
                return -1;
            }

            public void GetYearData(string name, string year)
            {
                var data = worksheet.Cells.Value as object[,];

                var startIndex = ConvertMonthYearToIndexInArray("Сентябрь", year);

            }
            
            public void AddPublisher(string name, bool pioner = false, bool pastor = false, bool ministryAssistant = false)
            {
                if (IsPublisherContainsInTable(name) == true)
                {
                    throw new Exception($"Возвещатель {name} уже есть в таблице.");
                }
                var data = worksheet.Cells.Value as object[,];
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
                using (var package = new ExcelPackage(new FileInfo(puthToFile)))
                {
                    var ws = package.Workbook.Worksheets[nameTable];
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

                    package.Save();
                }

            }

            public int GetMonthReports(string month, TypePublisher typePublisher, TypeGetPublisherResponceReport typeGetPublisherResponceReport, string year = "current")
            {
                if (year == "current")
                { year = DateTime.Now.Year.ToString(); }

                int currentMonth = ConvertMonthYearToIndexInArray(month, year);
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
                if (typeGetPublisherResponceReport != TypeGetPublisherResponceReport.count)
                { 
                    tgr = typeGetPublisherResponceReport; 
                }
                else
                {
                    tgr = TypeGetPublisherResponceReport.hour;
                }

                var data = worksheet.Cells.Value as object[,];

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

            public List<string> PublsherWithoutReport(string month, string year)
            {
                List<string> publishers = new List<string>();
                var data = worksheet.Cells.Value as object[,];

                int currentMonth = ConvertMonthYearToIndexInArray(month, year);

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
                var data = worksheet.Cells.Value as object[,];
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
            
            public List<List<object>> GetYearReportsPublisher(string name, string year)
            {
                var data = worksheet.Cells.Value as object[,];

                var rowIndex = GetRangePublisher(name);
                var columnIndex = ConvertMonthYearToIndexInArray("сентябрь", (Int32.Parse(year) - 1).ToString());

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
                var data = worksheet.Cells.Value as object[,];

                int startIndexColumn = ConvertMonthYearToIndexInArray("сентябрь", (Int32.Parse(year) - 1).ToString());
                int startIndexRow = GetRangePublisher(name).First() + 1;

                using (var package = new ExcelPackage(new FileInfo(puthToFile)))
                {
                    var ws = package.Workbook.Worksheets[nameTable];
                    List<string> asas = new List<string>();
                    int tempstartIColumn = startIndexColumn;
                    SymbolColumns = GetColumnSymbolAsStringArray(EndColumnData);
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
                    package.Save();
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

        public class ArchiveReports
        {
            private ExcelWorksheet worksheet;
            public readonly string puthToFile;
            public readonly string nameTable;

            public ArchiveReports(JwBookExcel excelData)
            {
                puthToFile = excelData.PuthToWorkbook;
                nameTable = excelData.NameTableArchiveReports;

                worksheet = excelData.ArchiveReportsWorkSheet;
            }

            public bool ContainsMonthReport(string month, string year)
            {
                var data = worksheet.Cells.Value as object[,];
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    if (data[i, 0] != null && data[i, 0].ToString().ToLower() == year.ToLower())
                    {
                        if (data[i, 1].ToString().ToLower() == month.ToLower())
                        {
                            return true;
                        }
                    }
                }
                return false;
            }

            public void UpdateArchive(object[,] values, string month, string year)
            {
                var arrColumnSym = GetColumnSymbolAsStringArray(((object[,])worksheet.Cells.Value).GetLength(1)); // 0 rows; 1 - columns
                var data = worksheet.Cells.Value as object[,];

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    if (data[i, 0] != null && data[i, 0].ToString().ToLower() == year.ToLower())
                    {
                        if (data[i, 1].ToString().ToLower() == month.ToLower())
                        {
                            i++; // Переведём на следующий индекс
                            using (var pacakge = new ExcelPackage(new FileInfo(puthToFile)))
                            {
                                var ws = pacakge.Workbook.Worksheets[nameTable];
                                for (int j = 0; j < values.GetLength(0); j++) // Rows - Publisher - Apioner, Pioner
                                {
                                    ws.Cells[$"D{i}"].Value = values[j, 0]; // Count
                                    ws.Cells[$"E{i}"].Value = values[j, 1]; // Publications
                                    ws.Cells[$"F{i}"].Value = values[j, 2]; // Videos
                                    ws.Cells[$"G{i}"].Value = values[j, 3]; // Hours
                                    ws.Cells[$"H{i}"].Value = values[j, 4]; // ReturnVisits
                                    ws.Cells[$"I{i}"].Value = values[j, 5]; // BiblStudy
                                    i++;
                                }
                                // // Change year
                                // // CHange month
                                pacakge.Save();
                            }
                            break;
                        }
                    }
                }


            }

            public void CreateArchive(object[,] values, string month, string year)
            {
                var data = worksheet.Cells.Value as object[,];

                int index = default;
                // Search where insert
                for (int i = 1; i < data.GetLength(0); i++)
                {
                    if(data[i, 0] != null && Int32.Parse(data[i,0].ToString()) == Int32.Parse(year))
                    {
                        if(ConvertMonthStringToInt32(data[i, 1].ToString()) > ConvertMonthStringToInt32(month))
                        {
                            index = i;
                            break;
                        }
                    }
                    else if(data[i, 0] != null && Int32.Parse(data[i, 0].ToString()) > Int32.Parse(year))
                    {
                        index = i;
                        break;
                    }
                    else
                    {
                        if(data[i, 0] == null)
                        {
                            index = i;
                            break;
                        }
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(puthToFile)))
                {
                    index++;
                    var ws = package.Workbook.Worksheets[nameTable];
                    ws.InsertRow(index, 3);

                    ws.Cells[$"C{index}"].Value = "Возвещатель";
                    ws.Cells[$"C{index + 1}"].Value = "Подсобный пионер";
                    ws.Cells[$"C{index + 2}"].Value = "Пионер";

                    for (int i = 0; i < values.GetLength(0); i++)
                    {
                        ws.Cells[$"A{index}"].Value = year;
                        ws.Cells[$"B{index}"].Value = month;

                        ws.Cells[$"D{index}"].Value = values[i, 0];
                        ws.Cells[$"E{index}"].Value = values[i, 1];
                        ws.Cells[$"F{index}"].Value = values[i, 2];
                        ws.Cells[$"G{index}"].Value = values[i, 3];
                        ws.Cells[$"H{index}"].Value = values[i, 4];
                        ws.Cells[$"I{index}"].Value = values[i, 5];
                        index++;
                    }
                    package.Save();
                }
            }

            public List<IList<object>> GetArchive(string[] selection, string typePublisher)
            {
                string fMonth = selection[0]; // first month
                string fYear = selection[1]; // first year
                string sMonth = selection[2]; // second month
                string sYear = selection[3]; // second year
                bool flagStartAdd = false;
                bool сancelLoop = false; // если подошел нужный месяц.

                List<IList<object>> archive = new List<IList<object>>();

                var data = worksheet.Cells.Value as object[,];
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    if (data[i, 0] != null && data[i, 0].ToString() == sYear)
                    {
                        if (data[i, 1] != null && data[i, 1].ToString() == sMonth)
                        {
                            сancelLoop = true;
                        }
                        else
                        {
                            if (сancelLoop == true)
                            {
                                break;
                            }
                        }
                    }
                    else
                    {
                        if (сancelLoop == true)
                            break;
                    }
                    if (data[i, 0] != null && data[i, 0].ToString() == fYear)
                    {
                        if (data[i, 1] != null && data[i, 1].ToString() == fMonth)
                        {
                            flagStartAdd = true;
                        }
                    }
                    if (flagStartAdd == true)
                    {
                        if (data[i, 0] != null)
                        {
                            List<object> month = new List<object>();
                            for (int j = 0; j < data.GetLength(1); j++)
                            {
                                month.Add(data[i, j]);
                            }
                            archive.Add(month);
                        }
                    }
                }

                if (typePublisher != "Все") // если не все возвещатели нужны, а только по назначениям
                {
                    archive.RemoveAll(r => r[2].ToString() != typePublisher);
                }
                return archive;
            }
        }

        public class NoActivePublishers
        {
            private ExcelWorksheet worksheet;
            public readonly string puthToFile;
            public readonly string nameTable;

            public NoActivePublishers(JwBookExcel excelData)
            {
                puthToFile = excelData.PuthToWorkbook;
                nameTable = excelData.NameTableNoActivePublisher;

                worksheet = excelData.NoActivePublisherWorkSheet;
            }

            public ObservableCollection<string[]> GetPublishers(string year)
            {
                var data = worksheet.Cells.Value as object[,];
                List<string[]> selectData = new List<string[]>();

                for (int i = 1; i < data.GetLength(0); i++)
                {
                    if (data[i, 1] != null && data[i, 0] != null)
                    {
                        string sheetDate = data[i, 0].ToString().Split('.').Last();
                        if (Int32.Parse(sheetDate) > Int32.Parse(year))
                        {
                            continue;
                        }
                        selectData.Add(new string[] { data[i, 0].ToString(), data[i, 1].ToString() });
                    }
                    else break;
                }
                var sortedData = from sd in selectData
                                 orderby sd[0].Split('.').Last() descending 
                                 select sd;
                ObservableCollection<string[]> returnSortedValue = new ObservableCollection<string[]>();
                foreach(var sd in sortedData)
                { 
                    returnSortedValue.Add(sd);
                }
                return returnSortedValue;
            }

            public bool IsPublisherNoActive(string name)
            {
                var data = worksheet.Cells.Value as object[,];
                
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    if (data[i, 0] != null && data[i, 0].ToString() == name)
                    {
                        return true;
                    }
                }
                return false;
            }

        }


    }
}
