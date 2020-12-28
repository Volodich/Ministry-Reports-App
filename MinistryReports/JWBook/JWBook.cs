using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using MinistryReports.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.CodeDom;
using System.Xml.Serialization;
using System.Windows.Documents;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using MinistryReports.Models.JWBook;
using System.Runtime;
using System.Diagnostics;
using MaterialDesignThemes.Wpf.Converters;
using Microsoft.Xaml.Behaviors.Media;
using System.Runtime.Remoting.Messaging;
using System.Collections.ObjectModel;
using Org.BouncyCastle.Bcpg;

namespace MinistryReports.JWBook
{
    /// <summary>
    /// Класс для работы с google sheets которая содержит данные служения возвещателей, архив подсчитанных отчётов и спосок неактивных.
    /// </summary>
    class JWBook
    {
        #region fields & property

        #region private
        static string[] Scopes = { SheetsService.Scope.Spreadsheets }; // Сущность для работы с таблицами. 

        static string ApplicationName = String.Empty; // Имя приложения. Задаёться один раз. Не влияет на роботу приложения.
        static string PuthToJsonFile { get; set; } // содержит строку пути JSON файла. Без файлы невозможны запросы.
        static string SpreadsheetId { get; set; } // Айди таблицы в целом.
        static string SheetDataMinistry { get; set; } // Название таблицы для работы с отчётами.
        static int? SheetIdMinistryReports { get; set; } // Айди таблицы - для заполнения данными через запрос.
        static string SheetArchiveReports { get; set; } // Название таблицы - содержит сумму отчётов за месяц служения собрания.
        static int? SheetIdArchiveReports { get; set; } // Айди таблицы - для заполнение данными через запрос.
        static string SheetNoActivityPublisher { get; set; }// // Название таблицы - содержит список неактивных возвещатлей и дату их последенего отчёта.
        static int? SheetIdNoActivePublisher { get; set; } // Айди таблицы  - джя заполнения данными через запрос.
        
        static SheetsService sheetsService;

        
        static int loopStepInDataPublisher = 6;
        
        private static int? MaxColumnCount { get; set; }
        private static int? MaxRowCount { get; set; }


        static readonly string sheetAlphabet = "A-B-C-D-E-F-G-H-I-J-K-L-M-N-O-P-Q-R-S-T-U-V-W-X-Y-Z-AA-AB-AC-AD-AE-AF-AG-AH-AI-AJ-AK-AL-AM-AN-AO-AP-AQ-AR-AS-AT-AU-AV-AW-AX-AY-AZ-BA-BB-BC-BD-BE-BF-BG-BH-BI-BJ-BK-BL-BM-BN-BO-BP-BQ-BR-BS-BT-BU-BV-BW-BX-BY-BZ-CA-CB-CC-CD-CE-CF-CG-CH-CI-CJ-CK-CL-CM-CN-CO-CP-CQ-CR-CS-CT-CU-CV-CW-CX-CY-CZ-DA-DB-DC-DD-DE-DF-DG-DH-DI-DJ-DK-DL-DM-DN-DO-DP-DQ-DR-DS-DT-DU-DV-DW-DX-DY-DZ-EA-EB-EC-ED-EE-EF-EG-EH-EI-EJ-EK-EL-EM-EN-EO-EP-EQ-ER-ES-ET-EU-EV-EW-EX-EY-EZ-FA-FB-FC-FD-FE-FF-FG-FH-FI-FJ-FK-FL-FM-FN-FO-FP-FQ-FR-FS-FT-FU-FV-FW-FX-FY-FZ-GA-GB-GC-GD-GE-GF-GG-GH-GI-GJ-GK-GL-GM-GN-GO-GP-GQ-GR-GS-GT-GU-GV-GW-GX-GY-GZ-HA-HB-HC-HD-HE-HF-HG-HH-HI-HJ-HK-HL-HM-HN-HO-HP-HQ-HR-HS-HT-HU-HV-HW-HX-HY-HZ-IA-IB-IC-ID-IE-IF-IG-IH-II-IJ-IK-IL-IM-IN-IO-IP-IQ-IR-IS-IT-IU-IV-IW-IX-IY-IZ-JA-JB-JC-JD-JE-JF-JG-JH-JI-JJ-JK-JL-JM-JN-JO-JP-JQ-JR-JS-JT-JU-JV-JW-JX-JY-JZ-KA-KB-KC-KD-KE-KF-KG-KH-KI-KJ-KK-KL-KM-KN-KO-KP-KQ-KR-KS-KT-KU-KV-KW-KX-KY-KZ-LA-LB-LC-LD-LE-LF-LG-LH-LI-LJ-LK-LL-LM-LN-LO-LP-LQ-LR-LS-LT-LU-LV-LW-LX-LY-LZ-MA-MB-MC-MD-ME-MF-MG-MH-MI-MJ-MK-ML-MM-MN-MO-MP-MQ-MR-MS-MT-MU-MV-MW-MX-MY-MZ-NA-NB-NC-ND-NE-NF-NG-NH-NI-NJ-NK-NL-NM-NN-NO-NP-NQ-NR-NS-NT-NU-NV-NW-NX-NY-NZ-OA-OB-OC-OD-OE-OF-OG-OH-OI-OJ-OK-OL-OM-ON-OO-OP-OQ-OR-OS-OT-OU-OV-OW-OX-OY-OZ-PA-PB-PC-PD-PE-PF-PG-PH-PI-PJ-PK-PL-PM-PN-PO-PP-PQ-PR-PS-PT-PU-PV-PW-PX-PY-PZ-QA-QB-QC-QD-QE-QF-QG-QH-QI-QJ-QK-QL-QM-QN-QO-QP-QQ-QR-QS-QT-QU-QV-QW-QX-QY-QZ-RA-RB-RC-RD-RE-RF-RG-RH-RI-RJ-RK-RL-RM-RN-RO-RP-RQ-RR-RS-RT-RU-RV-RW-RX-RY-RZ-SA-SB-SC-SD-SE-SF-SG-SH-SI-SJ-SK-SL-SM-SN-SO-SP-SQ-SR-SS-ST-SU-SV-SW-SX-SY-SZ-TA-TB-TC-TD-TE-TF-TG-TH-TI-TJ-TK-TL-TM-TN-TO-TP-TQ-TR-TS-TT-TU-TV-TW-TX-TY-TZ-UA-UB-UC-UD-UE-UF-UG-UH-UI-UJ-UK-UL-UM-UN-UO-UP-UQ-UR-US-UT-UU-UV-UW-UX-UY-UZ-VA-VB-VC-VD-VE-VF-VG-VH-VI-VJ-VK-VL-VM-VN-VO-VP-VQ-VR-VS-VT-VU-VV-VW-VX-VY-VZ-WA-WB-WC-WD-WE-WF-WG-WH-WI-WJ-WK-WL-WM-WN-WO-WP-WQ-WR-WS-WT-WU-WV-WW-WX-WY-WZ-XA-XB-XC-XD-XE-XF-XG-XH-XI-XJ-XK-XL-XM-XN-XO-XP-XQ-XR-XS-XT-XU-XV-XW-XX-XY-XZ-YA-YB-YC-YD-YE-YF-YG-YH-YI-YJ-YK-YL-YM-YN-YO-YP-YQ-YR-YS-YT-YU-YV-YW-YX-YY-YZ-ZA-ZB-ZC-ZD-ZE-ZF-ZG-ZH-ZI-ZJ-ZK-ZL-ZM-ZN-ZO-ZP-ZQ-ZR-ZS-ZT-ZU-ZV-ZW-ZX-ZY-ZZ-";
        static int startYear;
        
        private static ValueRange DataPublishers { get; set; } // вся информация, полученная из листа с отчётами возвещателей.
        private static ValueRange DataArchiveReports { get; set; } // Вся информация, полученная из листа с данными архива возвещателей.
        #endregion

        #region Public
        public static string СurrentYear { get => DateTime.Now.Year.ToString(); }
        public static int StartYear { get => startYear; } // Начальный год - по которому естоь отчёты.

        #endregion
        #endregion

        public JWBook() { }

        public JWBook(JWBookSettings settings)
        {
            ApplicationName = settings.ApplicationName;
            SpreadsheetId = settings.SpreadsheetId;
            SheetDataMinistry = settings.SheetNameDataMinistry;
            SheetIdMinistryReports = settings.SheetIdDataMinistry;
            SheetArchiveReports = settings.SheetNameArchiveReports;
            SheetIdArchiveReports = settings.SheetIdArchiveReport;
            SheetNoActivityPublisher = settings.SheetNameNoActivityPublishers;
            SheetIdNoActivePublisher = settings.SheetIdNoActivityPublisher;
            PuthToJsonFile = settings.PuthJsonFile;

            ConnectSpreadsheet(); // Попытка подключения к таблице. 
        }

        /// <summary>
        /// Подключение JSON файла.
        /// </summary>
        /// <returns>Результат запроса к Google Sheets</returns>
        private static bool ConnectFileJson()
        {
            GoogleCredential credential;
            using (var strean = new FileStream(PuthToJsonFile,
            FileMode.Open,
            FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(strean).CreateScoped(Scopes);
            }
            sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return true;
        }

        public static async Task<JWBook> JWbookFactory(JWBookSettings setting)
        {
            return await Task.Run(() => new JWBook(setting));
        }

        internal async Task<bool> ConnectSpreadsheetAsync()
        {
            return await Task<bool>.Run(() => this.ConnectSpreadsheet());
        }

        internal bool ConnectSpreadsheet()
        {
            try
            {
                ConnectFileJson(); // подключаем, а заодно и проверяем JSON файл.

                SpreadsheetsResource.GetRequest request = sheetsService.Spreadsheets.Get(SpreadsheetId);
                IList<Sheet> sheets = request.Execute().Sheets;
                // Получаем максимальное количество наших столбцов и строк в нужном листе
                // TODO: допилить ексепшин  если при подключении - первый лист не данные отчётов.
                MaxColumnCount = sheets[0].Properties.GridProperties.ColumnCount; // [0] - первый лист в списке таблиц - reports
                MaxRowCount = sheets[0].Properties.GridProperties.RowCount; // [0] - первый лист в списке таблиц - reports

                DataPublishers = DataMinistry.GetAllDataPublisher(SpreadsheetId);
                startYear = DataMinistry.GetStartYear();
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                //MessageBox.Show("DirectoryNotFoundException");
                throw new Exception("Неверно указан путь к JSON файлу.");
                //return false;
            }
            catch (Google.GoogleApiException)
            {
                //MessageBox.Show("Google.GoogleApiException");
                throw new Exception("Неверное указан SpreadSheetId Вашей таблици");
                //return false;
            }
            catch (System.Net.Http.HttpRequestException)
            {
                //MessageBox.Show("System.Net.Http.HttpRequestException");
                throw new Exception("Проверьте подключение к интернету. Ошибка подключения к Google таблицам");
                //return false;
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"Неизвестная ошшибка. Обратитесь к администратору. ExceptionInner: {ex.InnerException}. Exception Message: {ex.Message}");
                throw new Exception($"Неизвестная ошшибка. Обратитесь к администратору. ExceptionInner: {ex.InnerException}. Exception Message: {ex.Message}"); 
                //return false;
            }
            return true;
        }

        public static bool InsertRows(int startIndex, int endIndex, int sheetId)
        {
            //try
            //{
                InsertDimensionRequest insertRow = new InsertDimensionRequest();
                insertRow.Range = new DimensionRange()
                {
                    SheetId = sheetId,
                    Dimension = "ROWS",
                    StartIndex = startIndex,
                    EndIndex = endIndex
                };
                BatchUpdateSpreadsheetRequest r = new BatchUpdateSpreadsheetRequest()
                {
                    Requests = new List<Request>
                        {
                            new Request{ InsertDimension = insertRow },
                        }
                };

                BatchUpdateSpreadsheetResponse response1 = sheetsService.Spreadsheets.BatchUpdate(r, SpreadsheetId).Execute();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"Подробнее: {ex.Message}", "Ошибка");
            //    return false;
            //}
            return true;
        }
        
        /// <summary>
        /// Работа с отчётами возвещателей
        /// </summary>
        public class DataMinistry
        {
            /// <summary>
            /// Возвращает все данные конкретного возвещателя, что есть в таблице (за все года)
            /// </summary>
            /// <param name="namePublisher">Имя возвещателя. Формат - "Фамилия Имя"</param>
            /// <returns>Список объектов, хранящий информацию про конкретного возвещателя. Каждый объект - отдельная клетка в таблице. </returns>
            public static List<IList<object>> GetInfoPublisher(string namePublisher)
            {
                var getRangePublisher = JWBook.DataMinistry.GetRangePublisher(namePublisher, false);

                string[] rangePublisher = { getRangePublisher[0].ToString(), getRangePublisher[1].ToString() };
                string[] columnsABC = sheetAlphabet.Split('-');
                string range = $"{SheetDataMinistry}!A{rangePublisher[0]}:{rangePublisher[1]}{columnsABC[(int)MaxColumnCount]}";

                SpreadsheetsResource.ValuesResource.GetRequest request = sheetsService.Spreadsheets.Values.Get(SpreadsheetId, range);
                return (List<IList<object>>)request.Execute().Values;
            }

            /// <summary>
            /// Получает начальный год с которым полдьзователь работает.
            /// </summary>
            /// <returns></returns>
            public static  int GetStartYear()
            {
                string[] columnsABC = sheetAlphabet.Split('-');

                
                int startIndex = 2; // 0 - Фамилия имя; 1 - Публикации, видео, тд; 2 - сумма за несуществующий год.
                int monthCount = 12; //  Количество месяце в году. 
                int nextRow = 1; // Передвигаем указатель на следующею ячейку

                string range = $"{SheetDataMinistry}!{columnsABC[startIndex+monthCount+nextRow]}1";
                SpreadsheetsResource.ValuesResource.GetRequest requestAllPublisherData = sheetsService.Spreadsheets.Values.Get(SpreadsheetId, range);
                var tempData = (List<IList<object>>)requestAllPublisherData.Execute().Values;
                return Convert.ToInt32(tempData[0][0].ToString().Substring(0,4)); // Так как коллекция содержит всего один элемент - год.
            }

            /// <summary>
            /// Возвращает range возвещателя в таблице JWBook
            /// </summary>
            /// <param name="name">Имя возвещателя. Формат - "Фамилия Имя" </param>
            /// <returns>List(int) - содержит два значения. Первое - начальный range возвещателя в таблице. 
            /// Второе - последний rasnge возвещателяч в таблице. </returns>
            public static List<int> GetRangePublisher(string name, bool OfflineGet, ValueRange valueRange = null)
            {
                var range = $"{SheetDataMinistry}!A1:A{MaxRowCount}";

                int? firstRange = null;
                int? lastRange = null;
                List<IList<object>> resultValues;
                if (OfflineGet == false)
                {
                    SpreadsheetsResource.ValuesResource.GetRequest request = sheetsService.Spreadsheets.Values.Get(SpreadsheetId, range);
                    resultValues = (List<IList<object>>)request.Execute().Values;
                }
                else // if (OfflineGet == true)
                {
                        resultValues = (List<IList<object>>)valueRange.Values;
                }

                bool flag = false;
                for (int i = 0; i < resultValues.Count; i++)
                {
                    for (int j = 0; j < resultValues[i].Count; j++)
                    {
                        if ($"{name}" == resultValues[i][j].ToString())
                        {
                            firstRange = ++i;
                            lastRange = firstRange + 5;
                            flag = true;
                            break;
                        }
                    }
                    if (flag == true)
                    {
                        break;
                    }
                }
                if (firstRange != null && lastRange != null)
                {
                    List<int> PublisherRange = new List<int>();
                    PublisherRange.Add((int)firstRange);
                    PublisherRange.Add((int)lastRange);
                    return PublisherRange;
                }
                throw new NotSupportedException($"Не удалось найти {name} в таблице Google Sheet");
            }

            internal static ValueRange GetAllDataPublisher(string SpreadsheetId)
            {
                string[] columnsABC = sheetAlphabet.Split('-');

                string range = $"{SheetDataMinistry}!A2:{columnsABC[(int)MaxColumnCount + 1]}{MaxRowCount}";
                SpreadsheetsResource.ValuesResource.GetRequest requestAllPublisherData = sheetsService.Spreadsheets.Values.Get(SpreadsheetId, range);
                DataPublishers = requestAllPublisherData.Execute();
                return requestAllPublisherData.Execute();
            }

            internal static List<IList<object>> GetAllDataPublisher()
            {
                return (List<IList<object>>)DataPublishers.Values;
            }

            internal static bool CheckPublisher(string name)
            {
                var data = DataPublishers.Values;
                List<string> publishers = new List<string>();

                for (int i = 0; i < data.Count; i += 6)
                {
                    publishers.Add(data[i][0].ToString());
                }
                return publishers.Contains(name);
            }

            public static List<List<object>> GetOfflineDataPublisher(ValueRange data, string name, string year)
            {
                var datasPubl = (List<IList<object>>)data.Values;

                List<List<object>> datasPublishersConvert = new List<List<object>>();
                List<object> Publications = new List<object>();
                List<object> Videos = new List<object>();
                List<object> Hours = new List<object>();
                List<object> ReturnVisit = new List<object>();
                List<object> BiblStudy = new List<object>();
                List<object> Comment = new List<object>();

                int monthCount = 12;

                var rangePublisher = GetRangePublisher(name, true, data);
                rangePublisher[0] -= 1; // начинается информация на 1 рядок выше.
                int j = 0; // откуда начинать счёт - считается по формуле (yearConvert)
                #region yearConvert
                try
                {
                    int startIndex = 2;
                    int summColumn = 1;
                    int nextIndex = 1;
                    int startIndexInArray = startIndex + ((monthCount + summColumn) * (Convert.ToInt32(year) - StartYear)) + nextIndex; // ФОРМУЛА
                    /*
                     Массив даты пользователя состит из такого формата [I][J], где I - строки, J - колонки. На каждого возвещателя приходиться 6строк.
                     1 строка 1 колонка - Имя. 2 - описание единицы отчета (Публ., или видео, или часы...). 3 колонка - сумма за прошлый служебный год единицы
                     отчёта.
                     ПРИМЕР
                                    1(J) колонка    2 колонка       3 колонка        4 колонка       .....
                  1 (I)строка |Хромец Светлана	  |Публикации	|       1	       |    1       |
                     1 строка |д.о.	              |Видео		|       0	       |    0       |
                     1 строка |                   |Часы		    |       5	       |    5       |
                     1 строка |                   |Повт. Посещ.	|       5	       |    4       |
                     1 строка |                   |Библ. Изуч.	|       1	       |    0       |
                     1 строка |                   |Примечание	|       	       |    	    |

                    3 колонка - сумма за предыдущий служебный год данный из двеннадцати колонок. Таким образом сумма для последующих годов будет находиться через каждые
                    12 колонок. За 2018 год сумма находиться в колонке 3, за 2019 - 3+12=15 колонке. 
                    Формула, чтобы понять с какой колонки начинать одсчёи для извлечения данных из массива - 
                    Начало Отсчёта таблицы + ((Количество месяцев в служебном году + следующяя колонка ) * (текущий год - 2018)) + следующий индекс.
                     */
                    j = startIndexInArray;
                }
                catch (Exception)
                {
                    Console.WriteLine("Проверьте корректность параметра data");
                }
                #endregion

                for (int i = 0; i < monthCount; i++)
                {
                    Publications.Add(datasPubl[rangePublisher.First()][j]);
                    Videos.Add(datasPubl[rangePublisher.First() + 1][j]);
                    Hours.Add(datasPubl[rangePublisher.First() + 2][j]);
                    ReturnVisit.Add(datasPubl[rangePublisher.First() + 3][j]);
                    BiblStudy.Add(datasPubl[rangePublisher.First() + 4][j]);
                    Comment.Add(datasPubl[rangePublisher.First() + 5][j]);
                    j++;
                }

                datasPublishersConvert.Add(Publications);
                datasPublishersConvert.Add(Videos);
                datasPublishersConvert.Add(Hours);
                datasPublishersConvert.Add(ReturnVisit);
                datasPublishersConvert.Add(BiblStudy);
                datasPublishersConvert.Add(Comment);

                return datasPublishersConvert;
            }

            /// <summary>
            /// Возвращает заголовок столбца в google sheets. Например numberMonthColumn = 2 
            /// эквивалентно столбцу "B" в google sheets.
            /// </summary>
            /// <param name="numberMonthColumn">Номер столбца в таблице</param>
            /// <returns>Эквивалентное название столбца в таблице Google sheets</returns>
            public static string MonthConvert(int numberMonthColumn)
            {
                if (numberMonthColumn > 0 && numberMonthColumn < (int)MaxColumnCount)
                {
                    string[] columnABC = sheetAlphabet.Split('-');
                    return columnABC[numberMonthColumn];
                }
                else return String.Empty;
            }

            /// <summary>
            ///     
            /// </summary>
            /// <param name="month"></param>
            /// <param name="year"></param>
            /// <returns></returns>
            public static int MonthConvert(string month, string year)
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

                if (Int32.TryParse(year, out int resultYear))
                    return start + ((resultYear - StartYear) * (yearCount + sumColumn)) + intConvertMonth;
                else
                    return 0;
            }

            public static bool AddPblisher(object data, int startI)
            {
                var lists = data as List<IList<object>>;
                if (lists == null) return false;

                var bodyRequest = new ValueRange();
                bodyRequest.Values = lists;

                int startIndex = (startI * 6) + 2;
                int endIndex = (startI * 6) + 7;
                InsertRows(startIndex - 1, endIndex, SheetIdMinistryReports.Value);
                string range = $"{SheetDataMinistry}!A{startIndex}:{endIndex}";
                // Заносим имя возвещателя и доп информацию в таблицу
                var createPublisherInformationRequest = sheetsService.Spreadsheets.Values.Append(bodyRequest, SpreadsheetId, range);
                createPublisherInformationRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = createPublisherInformationRequest.Execute();
                // Создаём последнюю колонку чтобы небыло ексепшинов и была полная коллекция.
                string[] arrColumnName = sheetAlphabet.Split('-');
                string sumColumn = arrColumnName[MaxColumnCount.Value+1]; // Название последней колонки.
                string[] monthColumn = new string[2]; // 0 --- первая колонка - первый месяц служения. 1 --- последняя колонка - последний месяц служения.
                monthColumn[0] = arrColumnName[MaxColumnCount.Value - 11];
                monthColumn[1] = arrColumnName[MaxColumnCount.Value];

                List<IList<object>> sumMinistryYear = new List<IList<object>>();
 
                sumMinistryYear.Add(new List<object>() { $"=sum({SheetDataMinistry}!{monthColumn[0]}{startIndex}:{monthColumn[1]}{startIndex})" }); // Публикации.
                sumMinistryYear.Add(new List<object>() { $"=sum({SheetDataMinistry}!{monthColumn[0]}{startIndex + 1}:{monthColumn[1]}{startIndex + 1})" }); // Видео.
                sumMinistryYear.Add(new List<object>() { $"=sum({SheetDataMinistry}!{monthColumn[0]}{startIndex + 2}:{monthColumn[1]}{startIndex + 2})" }); // Часы.
                sumMinistryYear.Add(new List<object>() { $"=sum({SheetDataMinistry}!{monthColumn[0]}{startIndex + 3}:{monthColumn[1]}{startIndex + 3})" }); // Повт. посещения.
                sumMinistryYear.Add(new List<object>() { $"=sum({SheetDataMinistry}!{monthColumn[0]}{startIndex + 4}:{monthColumn[1]}{startIndex + 4})" }); // Изучения.
                sumMinistryYear.Add(new List<object>() { $"=sum({SheetDataMinistry}!{monthColumn[0]}{startIndex + 5}:{monthColumn[1]}{startIndex + 5})" }); // Публикации.

                range = $"{SheetDataMinistry}!{sumColumn}{startIndex}:{endIndex}";
                bodyRequest.Values = sumMinistryYear;

                var createSumYearMinistryRequest = sheetsService.Spreadsheets.Values.Append(bodyRequest, SpreadsheetId, range);
                createSumYearMinistryRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                appendResponse = createSumYearMinistryRequest.Execute();
                return true;
            }

            /// <summary>
            /// Получение необходимых данных про возвещателей собрания, месяц слежения, или теократический год
            /// </summary>
            public class GetPublisherReports
            {
                public class Month
                {

                    public int Publications(string month, TypePublisher typePublisher, string year = "default")
                    {
                        List<IList<object>> dataPublishers = null;
                        // online variant
                        //dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;
                        if (DataPublishers != null)
                        {
                            dataPublishers = (List<IList<object>>)DataPublishers.Values;
                        }
                        //ofline variant

                        if (year == "default")
                        { year = СurrentYear; }
                        int currentMonth = MonthConvert(month, year);
                        if (currentMonth <= 0)
                            throw new NotSupportedException("Неправильно указана дата! Проверьте еще раз. Возможно данных за такой месяц/год еще нет?");
                        int countPublicationsPioner = 0;
                        int countPublicationsAPioner = 0;
                        int countPublicationsPublisher = 0;

                        int loopStepPublisher = loopStepInDataPublisher;



                        for (int i = (int)TypeGetPublisherResponceReport.publications; i < dataPublishers.Count; i += 6)
                        {
                            if (dataPublishers.Count - 4 == i)
                            { break; }
                            if (loopStepPublisher + i >= dataPublishers.Count)
                            { break; }
                            if ((string)dataPublishers[i][currentMonth] != "")
                            {
                                if (Regex.IsMatch((string)dataPublishers[i + (loopStepPublisher - 1)][currentMonth], "\\bPP\\b"))
                                { // (string)dataPublishers[i + 2][0] == "PP" || 
                                    if (Int32.TryParse((string)dataPublishers[i][currentMonth], out int result))
                                        countPublicationsAPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else if ((string)dataPublishers[i + 2][0] == "ПИОНЕР")
                                {
                                    if (Int32.TryParse((string)dataPublishers[i][currentMonth], out int result))
                                        countPublicationsPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else
                                {
                                    if (Int32.TryParse((string)dataPublishers[i][currentMonth], out int result))
                                        countPublicationsPublisher += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                            }
                        }
                        switch (typePublisher)
                        {
                            case TypePublisher.Publisher:
                                return countPublicationsPublisher;
                            case TypePublisher.AuxiliaryPioneer:
                                return countPublicationsAPioner;
                            case TypePublisher.Pioneer:
                                return countPublicationsPioner;
                            case TypePublisher.InactivePublisher:
                                return -1;
                            // TODO: return values
                            default:
                                return -1;
                        }
                    }

                    public int Videos(string month, TypePublisher typePublisher, string year = "default")
                    {
                        List<IList<object>> dataPublishers = null;
                        // online variant
                        //dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;
                        if (DataPublishers != null)
                        {
                            dataPublishers = (List<IList<object>>)DataPublishers.Values;
                        }
                        //ofline variant

                        if (year == "default")
                        { year = СurrentYear; }
                        int currentMonth = MonthConvert(month, year);

                        int countVideosPioner = 0;
                        int countVideosAPioner = 0;
                        int countVideosPublisher = 0;

                        int loopStepPublisher = loopStepInDataPublisher;

                        for (int i = 0; i < dataPublishers.Count; i += 6)
                        {
                            if (i >= dataPublishers.Count - 5)
                            { break; }
                            if ((string)dataPublishers[i + 1][currentMonth] != "")
                            {
                                if (Regex.IsMatch((string)dataPublishers[i + (loopStepPublisher - 1)][currentMonth], "\\bPP\\b"))
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 1][currentMonth], out int result))
                                        countVideosAPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else if ((string)dataPublishers[i + 2][0] == "ПИОНЕР")
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 1][currentMonth], out int result))
                                        countVideosPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 1][currentMonth], out int result))
                                        countVideosPublisher += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                            }
                        }

                        switch (typePublisher)
                        {
                            case TypePublisher.Publisher:
                                return countVideosPublisher;
                            case TypePublisher.AuxiliaryPioneer:
                                return countVideosAPioner;
                            case TypePublisher.Pioneer:
                                return countVideosPioner;
                            case TypePublisher.InactivePublisher:
                                return -1;
                            // TODO: return values
                            default:
                                return -1;
                        }
                    }

                    public int Hours(string month, TypePublisher typePublisher, string year = "default")
                    {
                        List<IList<object>> dataPublishers = null;
                        // online variant
                        //dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;
                        if (DataPublishers != null)
                        {
                            dataPublishers = (List<IList<object>>)DataPublishers.Values;
                        }
                        //ofline variant

                        if (year == "default")
                        { year = СurrentYear; }
                        int currentMonth = MonthConvert(month, year);

                        int countHoursPioner = 0;
                        int countHoursAPioner = 0;
                        int countHoursPublisher = 0;

                        int loopStepPublisher = loopStepInDataPublisher;

                        for (int i = 0; i < dataPublishers.Count; i += 6)
                        {
                            if (i >= dataPublishers.Count - 5)
                            { break; }
                            if ((string)dataPublishers[i + 1][currentMonth] != "")
                            {
                                if (Regex.IsMatch((string)dataPublishers[i + (loopStepPublisher - 1)][currentMonth], "\\bPP\\b"))
                                { // (string)dataPublishers[i + (loopStepPublisher - 1)][currentMonth] == "PP"
                                    if (Int32.TryParse((string)dataPublishers[i + 2][currentMonth], out int result))
                                        countHoursAPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else if ((string)dataPublishers[i + 2][0] == "ПИОНЕР")
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 2][currentMonth], out int result))
                                        countHoursPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }

                                else
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 2][currentMonth], out int result))
                                        countHoursPublisher += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                            }
                        }
                        switch (typePublisher)
                        {
                            case TypePublisher.Publisher:
                                return countHoursPublisher;
                            case TypePublisher.AuxiliaryPioneer:
                                return countHoursAPioner;
                            case TypePublisher.Pioneer:
                                return countHoursPioner;
                            case TypePublisher.InactivePublisher:
                                return -1;
                            // TODO: return values
                            default:
                                return -1;
                        }
                    }

                    public int ReturnVisits(string month, TypePublisher typePublisher, string year = "default")
                    {
                        List<IList<object>> dataPublishers = null;
                        // online variant
                        //dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;
                        if (DataPublishers != null)
                        {
                            dataPublishers = (List<IList<object>>)DataPublishers.Values;
                        }
                        //ofline variant

                        if (year == "default")
                        { year = СurrentYear; }
                        int currentMonth = MonthConvert(month, year);

                        int countReturnVisitsPioner = 0;
                        int countReturnVisitsAPioner = 0;
                        int countreturnVisitsPublisher = 0;

                        int loopStepPublisher = loopStepInDataPublisher;

                        for (int i = 0; i < dataPublishers.Count; i += 6)
                        {
                            if (i >= dataPublishers.Count - 5)
                            { break; }
                            if ((string)dataPublishers[i + 1][currentMonth] != "")
                            {
                                if (Regex.IsMatch((string)dataPublishers[i + (loopStepPublisher - 1)][currentMonth], "\\bPP\\b"))
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 3][currentMonth], out int result))
                                        countReturnVisitsAPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else if ((string)dataPublishers[i + 2][0] == "ПИОНЕР")
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 3][currentMonth], out int result))
                                        countReturnVisitsPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 3][currentMonth], out int result))
                                        countreturnVisitsPublisher += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                            }
                        }
                        switch (typePublisher)
                        {
                            case TypePublisher.Publisher:
                                return countreturnVisitsPublisher;
                            case TypePublisher.AuxiliaryPioneer:
                                return countReturnVisitsAPioner;
                            case TypePublisher.Pioneer:
                                return countReturnVisitsPioner;
                            case TypePublisher.InactivePublisher:
                                return -1;
                            // TODO: return values
                            default:
                                return -1;
                        }
                    }

                    public int BiblStudy(string month, TypePublisher typePublisher, string year = "default")
                    {
                        List<IList<object>> dataPublishers = null;
                        // online variant
                        //dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;
                        if (DataPublishers != null)
                        {
                            dataPublishers = (List<IList<object>>)DataPublishers.Values;
                        }
                        //ofline variant

                        if (year == "default")
                        { year = СurrentYear; }
                        int currentMonth = MonthConvert(month, year);

                        int countVideosPioner = 0;
                        int countVideosAPioner = 0;
                        int countVideosPublisher = 0;

                        int loopStepPublisher = loopStepInDataPublisher;

                        for (int i = 0; i < dataPublishers.Count; i += 6)
                        {
                            if (i >= dataPublishers.Count - 5)
                            { break; }
                            if ((string)dataPublishers[i + 1][currentMonth] != "")
                            {
                                if (Regex.IsMatch((string)dataPublishers[i + (loopStepPublisher - 1)][currentMonth], "\\bPP\\b"))
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 4][currentMonth], out int result))
                                        countVideosAPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else if ((string)dataPublishers[i + 2][0] == "ПИОНЕР")
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 4][currentMonth], out int result))
                                        countVideosPioner += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 4][currentMonth], out int result))
                                        countVideosPublisher += result;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                            }
                        }
                        switch (typePublisher)
                        {
                            case TypePublisher.Publisher:
                                return countVideosPublisher;
                            case TypePublisher.AuxiliaryPioneer:
                                return countVideosAPioner;
                            case TypePublisher.Pioneer:
                                return countVideosPioner;
                            case TypePublisher.InactivePublisher:
                                return -1;
                            // TODO: return values
                            default:
                                return -1;
                        }
                    }

                    /// <summary>
                    /// Количество возвещателей сдавших отчёт
                    /// </summary>
                    /// <param name="month">месяц</param>
                    /// <param name="typePublisher">тип возвещателя</param>
                    /// <returns>количество возвещателей, сдавших отчёт</returns>
                    public int CountReports(string month, TypePublisher typePublisher, string year = "default")
                    {
                        List<IList<object>> dataPublishers = null;
                        // online variant
                        //dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;
                        if (DataPublishers != null)
                        {
                            dataPublishers = (List<IList<object>>)DataPublishers.Values;
                        }
                        //ofline variant

                        if (year == "default")
                        { year = СurrentYear; }
                        int currentMonth = MonthConvert(month, year);

                        int countCountReportsPioner = 0;
                        int countCountReportsAPioner = 0;
                        int countCountReportsPublisher = 0;

                        int loopStepPublisher = loopStepInDataPublisher;

                        for (int i = 0; i < dataPublishers.Count; i += 6)
                        {
                            if (i >= dataPublishers.Count - 5)
                            { break; }
                            if ((string)dataPublishers[i + 1][currentMonth] != "")
                            {
                                if (Regex.IsMatch((string)dataPublishers[i + (loopStepPublisher - 1)][currentMonth], "\\bPP\\b"))
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 2][currentMonth], out int result))
                                        countCountReportsAPioner++;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                                else if ((string)dataPublishers[i + 2][0] == "ПИОНЕР")
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 2][currentMonth], out int result))
                                        countCountReportsPioner++;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");

                                }
                                else
                                {
                                    if (Int32.TryParse((string)dataPublishers[i + 2][currentMonth], out int result))
                                        countCountReportsPublisher++;
                                    else throw new NotSupportedException($"Неверные данные в таблице, смотрите ячейку возвещателя {dataPublishers[i][0].ToString()} за {month} {year}");
                                }
                            }
                        }
                        switch (typePublisher)
                        {
                            case TypePublisher.Publisher:
                                return countCountReportsPublisher;
                            case TypePublisher.AuxiliaryPioneer:
                                return countCountReportsAPioner;
                            case TypePublisher.Pioneer:
                                return countCountReportsPioner;
                            case TypePublisher.InactivePublisher:
                                return -1;
                            // TODO: return values
                            default:
                                return -1;
                        }
                    }

                    /// <summary>
                    /// Количество активных возвещателей
                    /// </summary>
                    /// <param name="month">месяц, за который нужно получить ответ</param>
                    /// <returns>число активных возвещателей в указанный месяц.</returns>
                    public int CountActivePublisher(string month, string year = "default")
                    {
                        List<IList<object>> dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;

                        if (year == "default")
                        { year = СurrentYear; }
                        year = СurrentYear;
                        int currentMonth = MonthConvert(month, year);

                        int countActivePublisher = 0;
                        int tempCurrentMonth = 0;
                        int loopStepPublisher = loopStepInDataPublisher;

                        for (int i = 0; i < dataPublishers.Count; i += 6)
                        {
                            if (i >= dataPublishers.Count - 5)
                            { break; }
                            if ((string)dataPublishers[i + 2][currentMonth] == "" || (string)dataPublishers[i + 2][currentMonth] == "0")
                            {
                                tempCurrentMonth = currentMonth;
                                if (currentMonth > 6) // проверем, есть ли нужное количество отчётов.
                                {
                                    for (int j = 0; j < 6; j++)
                                    {
                                        tempCurrentMonth--;
                                        if (tempCurrentMonth == ((Int32.Parse(year) - 2019) * 13) + 2) // обходим сумму за год (у неактивных 0) - число всеравно есть в колонке 
                                            continue;
                                        if ((string)dataPublishers[i + 2][tempCurrentMonth] != "" && (string)dataPublishers[i + 2][tempCurrentMonth] != "0")
                                        {
                                            countActivePublisher++; 
                                            break;
                                        }
                                    }
                                }
                            }
                            else countActivePublisher++;
                        }
                        return countActivePublisher;
                    }

                    public List<string> PublisherWithoutReport(string month, string year)
                    {

                        List<string> publishers = new List<string>(); // возвещатели без отчёта
                        List<IList<object>> dataPublishers = null;
                        // online variant
                        //dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;
                        if (DataPublishers != null)
                        {
                            dataPublishers = (List<IList<object>>)DataPublishers.Values;
                        }
                        //ofline variant

                        int currentMonth = MonthConvert(month, year);

                        int loopStepPublisher = loopStepInDataPublisher;

                        for (int i = 0; i < dataPublishers.Count; i += 6)
                        {
                            if (i >= dataPublishers.Count - 5)
                            { break; }
                            if ((string)dataPublishers[i + 1][currentMonth] == "")
                            {
                                publishers.Add(dataPublishers[i][0].ToString());
                            }
                        }
                        return publishers;
                    }

                    public static  void SetNRPublisher(List<string> publishersNR, string month, string year) 
                    {
                        if (publishersNR == null) // Если нал - выходим из метода вооще.
                            goto exitLabel;

                        IList<IList<object>> values = new List<IList<object>>();
                        values.Add(new List<object>() { "NR" });
                        ValueRange bodyRequest = new ValueRange();
                        bodyRequest.Values = values;

                        foreach (var name in publishersNR)
                        {
                            var rangePublisher = GetRangePublisher(name, true, DataPublishers);
                            string column = MonthConvert(MonthConvert(month, year));
                            string range = $"{SheetDataMinistry}!{column}{rangePublisher.Last()+1}";
                            // Update Date
                            SpreadsheetsResource.ValuesResource.UpdateRequest update = sheetsService.Spreadsheets.Values.Update(bodyRequest, SpreadsheetId, range);
                            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                            update.Execute();
                        }
                       
                    exitLabel:;
                    }

                    public void ClearNRPublisher(string name) {}

                    /// <summary>
                    /// Возвращяет имена неактивных возвещателей колекцией string.
                    /// </summary>
                    /// <param name="month"> месяц, за который идёт запрос. По умолчанию - сентябрь.</param>
                    /// <param name="year">год, за который идёт запрос. По умолчанию - 2020.</param>
                    /// <returns>Колеекиця имён (string)</returns>
                    public List<string> InActivePublishers(string month, string year = "default")
                    {
                        List<IList<object>> dataPublishers = (List<IList<object>>)GetAllDataPublisher(SpreadsheetId).Values;

                        int currentMonth = 0;
                        if (month == "сентябрь" && year == "2020") // если параметры по умалчанию не сипользуються - береться текущее время
                        {
                            //currentMonth = MonthConvert(DateTime.Now.Month, DateTime.Now.Year);
                        }
                        else currentMonth = MonthConvert(month, year);

                        int tempCurrentMonth = 0;
                        List<string> inActivePublishers = new List<string>();

                        for (int i = 0; i < dataPublishers.Count; i += 6)
                        {
                            if (i >= dataPublishers.Count - 5)
                            {
                                break;
                            }
                            if ((string)dataPublishers[i + 2][currentMonth] == "" || (string)dataPublishers[i + 2][currentMonth] == "0")
                            {
                                tempCurrentMonth = currentMonth;
                                if (currentMonth > 6) // проверем, есть ли нужное количество отчётов.
                                {
                                    for (int j = 0; j < 7; j++)
                                    {
                                        // j < 7 потому что при итерации цикла, в феврале попадает так, что j >= 5 пропускается из-за того, что передним идёт колонка суммы, 
                                        //  и она операторам continue проходит дальше по циклу, миную добавления в коллекцию неактивного возвещателя.
                                        tempCurrentMonth--;
                                        if (tempCurrentMonth == ((Int32.Parse(year) - 2019) * 13) + 2) // обходим сумму за год (у неактивных 0) - число всеравно есть в колонке 
                                            continue;
                                        if (j >= 5) // про это.
                                        {
                                            inActivePublishers.Add((string)dataPublishers[i][0]);
                                            break;
                                        }
                                        if ((string)dataPublishers[i + 2][tempCurrentMonth] != "" && (string)dataPublishers[i + 2][tempCurrentMonth] != "0")
                                        {
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        return inActivePublishers;
                    }
                }
            }

            public class UpdatePublisherReports
            {
                public static bool UpdateYearData(object valueRange, string name, string year)
                {
                    var lists = valueRange as List<IList<object>>;
                    if (lists == null) return false;

                    var bodyRequest = new ValueRange();
                    bodyRequest.Values = lists;

                    var rangePubl = GetRangePublisher(name, true, DataPublishers);
                    string startColumn = MonthConvert(MonthConvert("Сентябрь", year));
                    string endColumn = MonthConvert(MonthConvert("Август", "2019"));

                    string range = $"{SheetDataMinistry}!{startColumn}{rangePubl.First() + 1}:{endColumn}{rangePubl.Last() + 1}";
                    SpreadsheetsResource.ValuesResource.UpdateRequest update = sheetsService.Spreadsheets.Values.Update(bodyRequest, SpreadsheetId, range);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    update.Execute();
                    return true;
                }
            }
        }       
        
        /// <summary>
        /// Работа с архивом подсчитанных отчётов собрания
        /// </summary>
        public class MeetReport
        {

            public MeetReport()
            {
                SpreadsheetsResource.ValuesResource.GetRequest request = sheetsService.Spreadsheets.Values.Get(SpreadsheetId, $"{SheetArchiveReports}!A:I");
                DataArchiveReports = request.Execute();
            }

            public static bool CheckMonthReport(DateTime date)
            {
                if(DataArchiveReports == null)
                { MeetReport mr = new MeetReport(); }
                
                var datas = DataArchiveReports.Values;

                string year = date.Year.ToString();
                string month = string.Empty;

                switch (date.Month)
                {
                    case 1:
                        month = "Январь";
                        break;
                    case 2:
                        month = "Февраль";
                        break;
                    case 3:
                        month = "Март";
                        break;
                    case 4:
                        month = "Апрель";
                        break;
                    case 5:
                        month = "Май";
                        break;
                    case 6:
                        month = "Июнь";
                        break;
                    case 7:
                        month = "Июль";
                        break;
                    case 8:
                        month = "Август";
                        break;
                    case 9:
                        month = "Сентябрь";
                        break;
                    case 10:
                        month = "Октябрь";
                        break;
                    case 11:
                        month = "Ноябрь";
                        break;
                    case 12:
                        month = "Декабрь";
                        break;
                    default:
                        throw new Exception("Не удалось определить месяц!");
                }
                try
                {
                    for (int i = 0; i < datas.Count; i++)  // TODO: LINQ
                    {
                        if ((string)datas[i][0] == year)
                            if ((string)datas[i][1] == month)
                            {
                                return true;
                            }
                    }
                }
                catch (Exception)
                {
                    throw new Exception($"Ошибка чтения данных с листа: {SheetArchiveReports}");
                }
                return false;
            }
            
            public bool UpdateArchive(object values, string month, string year)
            {
                var lists = values as List<IList<object>>;
                if (lists == null) return false;

                var valueRange = new ValueRange();
                valueRange.Values = lists;

                var datas = DataArchiveReports.Values;
                int startIndex = default;
                try
                {
                    for (int i = 0; i < datas.Count; i++)  // TODO: LINQ
                    {
                        if ((string)datas[i][0] == year)
                            if ((string)datas[i][1] == month)
                            {
                                startIndex = i;
                                break;
                            }
                    }
                }
                catch(Exception ex) // Если произойдёт ошибка при обработке данных.
                {
                    throw ex;
                }
                var range = $"{SheetArchiveReports}!A{startIndex + 1}:I{startIndex + 3}";

                var appendRequest = sheetsService.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute();

                return true;
            }

            public bool CreateArchive(object values, string month, string year)
            {
                var lists = values as List<IList<object>>;
                if (lists == null) return false;

                var valueRange = new ValueRange();
                valueRange.Values = lists;

                var datas = DataArchiveReports.Values;
                int startIndex = datas.Count;

                try
                {
                    for (int i = 1; i < datas.Count; i++)
                    {
                        if (Int32.Parse(year) < Int32.Parse((string)datas[i][0]))
                        {
                            startIndex = i;
                            InsertRows(startIndex, startIndex + 3, SheetIdArchiveReports.Value);
                            break;
                        }
                    }
                }
                catch (Exception)
                {
                    throw new Exception($"Проверьте содержимое листа: {SheetArchiveReports}");
                }

                var range = $"{SheetArchiveReports}!A{startIndex}:I{startIndex + 2}";

                var appendRequest = sheetsService.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute();

                return true;
            }

            /// <summary>
            /// Возвращает данные о служении собрания за определенное время, выбранное пользователем.
            /// </summary>
            /// <param name="selection">array string. 4 elements. 1 - first month, 2 - first year, 3 - second month, 4 second year</param>
            /// <param name="typePublisher">Type of publisher</param>
            public List<IList<object>> GetArchive(string[] selection, string typePublisher)
            {
                string fMonth = selection[0]; // first month
                string fYear = selection[1]; // first year
                string sMonth = selection[2]; // second month
                string sYear = selection[3]; // second year
                bool flagStartAdd = false;
                bool сancelLoop = false; // если подошел нужный месяц.

                List<IList<object>> response = new List<IList<object>>();

                List<IList<object>> datasArchive = (List<IList<object>>)DataArchiveReports.Values;
                for (int i = 0; i < datasArchive.Count; i++)
                {
                    if ((string)datasArchive[i][0] == sYear) // если совпал год
                    {
                        if ((string)datasArchive[i][1] == sMonth)
                        {
                            сancelLoop = true;
                        }
                        else
                        {
                            if (сancelLoop == true)
                                break;
                        }
                    }
                    else
                    {
                        if (сancelLoop == true)
                            break;
                    }
                    if ((string)datasArchive[i][0] == fYear)
                        if ((string)datasArchive[i][1] == fMonth)
                            flagStartAdd = true;
                    if (flagStartAdd)
                        response.Add(datasArchive[i]);
                }

                if (typePublisher != "Все") // если не все возвещатели нужны, а только по назначениям
                    response.RemoveAll(r => r[2].ToString() != typePublisher);
                return response;
            }

            

            public bool ItContains(string month, string year)
            {
                var values = DataArchiveReports.Values;
                foreach (var value in values) // TODO: LINQ
                {
                    if (value.Contains(year))
                        if (value.Contains(month))
                            return true;
                }
                return false;
            }
        }

        public class NoActivityPublisher
        {
            private ValueRange DataNoActivityPublisher { get; set; }

            public NoActivityPublisher()
            {
                RefreshData(); // получаем данные таблицы.
            }

            /// <summary>
            /// Получаем список неактивных возвещателей.
            /// </summary>
            /// <returns>Коллекцию строковых массивов. Элементы массива: [0] - Дата, когда возвещатель стал нективным; [1] - Имя возвещателя;</returns>
            public ObservableCollection<string[]> GetRequest(string year = "current") // [0] - Дата, когда возвещатель стал нективным [1] - Имя возвещателя.
            {
                if (year == "current")
                {
                    year = DateTime.Now.Year.ToString();
                }

                var data = (List<IList<object>>)DataNoActivityPublisher.Values;

                ObservableCollection<string[]> selectData = new ObservableCollection<string[]>();

                int firstRow = 0;
                int row = 0;
                foreach(var publisher in data)
                {
                    if(firstRow == row)
                    {
                        row++;
                        continue;
                    }
                    if( Convert.ToDateTime(publisher[0]).Year > Int32.Parse(year)) // Если возвещатель в таблице стал неактивным позже указанного года - остановить цикл.
                    {
                        break;
                    }
                    selectData.Add(new string[] {publisher[0].ToString(), publisher[1].ToString()});
                }
                return selectData; 
            }

            /// <summary>
            /// Изменить данные относительно неактивного возвещателя.
            /// </summary>
            public void UpdateRequest(ObservableCollection<string[]> updateData) 
            {
                // Create Body Request
                IList<IList<object>> dataRequest = new List<IList<object>>();
                foreach(var u in updateData)
                {
                    dataRequest.Add(new List<object> {u[0], u[1]}); // Элементы массива: [0] - Дата, когда возвещатель стал неактивным; [1] - Имя возвещателя.
                }

                ValueRange bodyRequest = new ValueRange();
                bodyRequest.Values = dataRequest;

                string range = $"{SheetNoActivityPublisher}!A2:B{dataRequest.Count+1}"; 
                SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest = sheetsService.Spreadsheets.Values.Update(bodyRequest, SpreadsheetId, range); // Column A - Date; Column B - Name;
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var response = updateRequest.Execute();
            }

            /// <summary>
            /// Возвещатель перешёл в другое собрание или снова стал активным.
            /// </summary>
            public void DeleteRequest(string name, string date)  
            {
                var data = (List<IList<Object>>)DataNoActivityPublisher.Values;

                int deleteIndex = default;

                for (int i = 0; i < data.Count; i++)
                {
                    if(data[i][1].ToString() == name && data[i][0].ToString() == date) // Если совпало имя, проверяем дату, чтобы наверняка.
                    {
                        deleteIndex = i;
                        break;
                    }    
                }

                if(deleteIndex == default) // Если не нашли совпадение.
                {
                    throw new KeyNotFoundException($"Ошибка удаления. Не удалось найти возвещателя в списке {SheetNoActivityPublisher} с таким именем: {name}");
                }

                Request request = new Request()
                {
                    DeleteDimension = new DeleteDimensionRequest()
                    {
                        Range = new DimensionRange()
                        {
                            SheetId = SheetIdNoActivePublisher.Value,
                            Dimension = "ROWS",
                            StartIndex = deleteIndex,
                            EndIndex = deleteIndex+1
                        }
                    }
                };

                BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
                requestBody.Requests = new List<Request>() { request };

                SpreadsheetsResource.BatchUpdateRequest deletionRequest = new SpreadsheetsResource.BatchUpdateRequest(sheetsService, requestBody, SpreadsheetId);
                deletionRequest.Execute();
            }

            /// <summary>
            /// Возвещатель стал неактивным.
            /// </summary>
            public void CreateRequest(string name, string date) 
            {
                var data = (List<IList<object>>)DataNoActivityPublisher.Values;
                
                List<string> namesNoActivityPublisher = new List<string>();

                foreach(var d in data)
                {
                    namesNoActivityPublisher.Add(d[1].ToString());
                }
                namesNoActivityPublisher.Add(name); // Добавили нового возвещателя. По-умолче=анию добавили в конец.
                namesNoActivityPublisher.Sort(); // Отсортировали возвещателей по алфавиту.
                var indexPublisher = namesNoActivityPublisher.IndexOf(name);

                InsertRows(indexPublisher, indexPublisher, SheetIdNoActivePublisher.Value);

                List<IList<object>> valuesRequest = new List<IList<object>> { new List<object>() { date, name} };
                ValueRange bodyRequest = new ValueRange();
                bodyRequest.Values = valuesRequest;

                string range = $"{SheetIdNoActivePublisher}!A{indexPublisher}:B{indexPublisher}";

                var appendRequest = sheetsService.Spreadsheets.Values.Append(bodyRequest, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute();
            }

            /// <summary>
            /// Проверка, возвещатель являеться ли неактивным?
            /// </summary>
            /// <returns></returns>
            public bool CheckPublisher(string name, string date) 
            {
                var data = DataNoActivityPublisher.Values;
                foreach(var d  in data)
                {
                    if(d[0].ToString() == date && d[1].ToString() == name)
                    {
                        return true;
                    }
                }
                return false; 
            }

            /// <summary>
            /// Обновляет данные о неактивных возвещателях.
            /// </summary>
            public void RefreshData() 
            {
                SpreadsheetsResource.ValuesResource.GetRequest request = sheetsService.Spreadsheets.Values.Get(SpreadsheetId, $"{SheetNoActivityPublisher}!A:B"); // Column A - Date; Column B - Name;
                DataNoActivityPublisher = request.Execute();
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
        biblStudy
    }
}
