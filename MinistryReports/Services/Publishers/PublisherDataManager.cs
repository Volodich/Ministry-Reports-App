using System.Collections.Generic;
using System.IO;
using System.Linq;
using MinistryReports.Models.S21;
using OfficeOpenXml;

namespace MinistryReports.Services.Publishers
{
    public interface IPublisherDataManager
    {
        void CheckConnect();
        IEnumerable<PublishersRange> GetPublishers();
        void AddPublisher(PublishersRange publisher);
    }
    public class PublisherDataManager : IPublisherDataManager
    {
        private ExcelWorksheet _publishersWorksheet;
        private ExcelPackage _package;

        private readonly string _pathToFile;
        static string NameTable => ApplicationConfig.PublisherInfo.TableName;

        public PublisherDataManager(S21Settings settings)
        {
            if (settings != null)
            {
                _pathToFile = settings.PuthToExcelDbFile;
            }

            CheckConnect();
        }

        public void CheckConnect()
        {
            FileInfo fileInfo = new FileInfo(_pathToFile);

            if (fileInfo.Exists)
            {
                _package = new ExcelPackage(fileInfo);
                ExcelWorksheet excelWorksheets = _package.Workbook.Worksheets[NameTable];
                if (excelWorksheets != null)
                {
                    this._publishersWorksheet = excelWorksheets;
                }
                throw new FileLoadException($"Таблица не содержит данных возвещателей. Проверьте правильный ли Вы выбрали excel файл. Либо есть ли в нужном файле лист {NameTable}.");
            }
            throw new FileNotFoundException($"Невозможно найти файл по такому расположению: {_pathToFile}");
        }

        public IEnumerable<PublishersRange> GetPublishers()
        {
            var dataPublishers = (object[,])_publishersWorksheet.Cells.Value;
            List<PublishersRange> publishers = new List<PublishersRange>(dataPublishers.GetLength(0));

            for (int i = 1; i < dataPublishers.GetLength(0); i++)
            {
                publishers.Add(new PublishersRange()
                {
                    Name = dataPublishers[i, 0]?.ToString(),
                    DateBirth = dataPublishers[i, 1]?.ToString() ?? "",
                    BuptismDate = dataPublishers[i, 2]?.ToString() ?? "",
                    Adress = dataPublishers[i, 3]?.ToString() ?? "",
                    Gender = dataPublishers[i, 4]?.ToString() ?? "",
                    Pioner = dataPublishers[i, 5]?.ToString() ?? "",
                    Mobile1 = dataPublishers[i, 6]?.ToString() ?? "",
                    Mobile2 = dataPublishers[i, 7]?.ToString() ?? "",
                    Appointment = dataPublishers[i, 9]?.ToString() ?? "",
                    Group = dataPublishers[i, 10]?.ToString() ?? "",
                    CountId = i
                });
            }
            return publishers;
        }

        // TODO: refact: сделать бул возврат. Если н удалось добавить вернуть фолс и сделать запись в лог журнал.
        public void AddPublisher(PublishersRange publisher)
        {
            var startIndex = GetPublisherIndex(publisher);

            using (var p = new ExcelPackage(_package.File))
            {
                ExcelWorksheet ws = _package.Workbook.Worksheets[NameTable];
                ws.InsertRow(startIndex, 1);
                ws.Cells[$"A{startIndex}"].Value = publisher.Name; // Имя Фамилия возвещателя
                ws.Cells[$"B{startIndex}"].Value = publisher.DateBirth; // Дата рождения
                ws.Cells[$"C{startIndex}"].Value = publisher.BuptismDate; // Дата крещения
                ws.Cells[$"D{startIndex}"].Value = publisher.Adress; // Адресс
                ws.Cells[$"E{startIndex}"].Value = publisher.Gender; // Пол
                ws.Cells[$"F{startIndex}"].Value = publisher.Pioner; // Пионер
                ws.Cells[$"G{startIndex}"].Value = publisher.Mobile1; // Номер телефона 1
                ws.Cells[$"H{startIndex}"].Value = publisher.Mobile2; // Номер телефона 2
                ws.Cells[$"I{startIndex}"].Value = publisher.Appointment; // Назначение (братья)

                _package.Save();
            }
        }

        private int GetPublisherIndex(PublishersRange publisher)
        {
            var publishersInfo = GetPublishers().ToList();
            publishersInfo.Add(publisher);

            return publishersInfo.OrderBy(p => p.Name).ToList().IndexOf(publisher);
        }
    }
}
